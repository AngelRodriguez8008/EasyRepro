﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Threading;
using System.Web;
using OtpNet;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class WebClient : BrowserPage
    {
        public List<ICommandResult> CommandResults => Browser.CommandResults;
        public Guid ClientSessionId;

        public WebClient(BrowserOptions options)
        {
            Browser = new InteractiveBrowser(options);
            OnlineDomains = Constants.Xrm.XrmDomains;
            ClientSessionId = Guid.NewGuid();
        }

        internal BrowserCommandOptions GetOptions(string commandName)
        {
            return new BrowserCommandOptions(Constants.DefaultTraceSource,
                commandName,
                Constants.DefaultRetryAttempts,
                Constants.DefaultRetryDelay,
                null,
                true,
                typeof(NoSuchElementException), typeof(StaleElementReferenceException));
        }

        internal BrowserCommandResult<bool> InitializeModes(bool onlineLoginPath = false)
        {
            return Execute(GetOptions("Initialize Unified Interface Modes"), driver =>
            {
                var uri = driver.Url;
                var queryParams = "";

                if (Browser.Options.UCITestMode) queryParams += "&flags=testmode=true,easyreproautomation=true";
                if (Browser.Options.UCIPerformanceMode) queryParams += "&perf=true";

                if (!string.IsNullOrEmpty(queryParams) && !uri.Contains(queryParams))
                {
                    var testModeUri = uri + queryParams;

                    driver.Navigate().GoToUrl(testModeUri);

                    driver.WaitForPageToLoad();

                    if (!onlineLoginPath)
                    {
                        driver.WaitForTransaction();
                    }
                }

                return true;
            });
        }


        public string[] OnlineDomains { get; set; }

        #region Login
        internal BrowserCommandResult<LoginResult> Login(Uri uri)
        {
            var username = Browser.Options.Credentials.Username;
            if (username == null)
                return PassThroughLogin(uri);

            var password = Browser.Options.Credentials.Password;
            return Login(uri, username, password);
        }

        internal BrowserCommandResult<LoginResult> Login(Uri orgUri, SecureString username, SecureString password, SecureString mfaSecrectKey = null, Action<LoginRedirectEventArgs> redirectAction = null)
        {
            return Execute(GetOptions("Login"), Login, orgUri, username, password, mfaSecrectKey, redirectAction);
        }

        private LoginResult Login(IWebDriver driver, Uri uri, SecureString username, SecureString password, SecureString mfaSecrectKey = null, Action<LoginRedirectEventArgs> redirectAction = null)
        {
            bool online = !(OnlineDomains != null && !OnlineDomains.Any(d => uri.Host.EndsWith(d)));
            driver.Navigate().GoToUrl(uri);

            if (!online)
                return LoginResult.Success;

            if (driver.IsVisible(By.Id("use_another_account_link")))
                driver.ClickWhenAvailable(By.Id("use_another_account_link"));

            bool waitingForOtc = false;
            bool success = EnterUserName(driver, username);
            if (!success)
            {
                var isUserAlreadyLogged = IsUserAlreadyLogged(driver);
                if (isUserAlreadyLogged)
                {
                    SwitchToDefaultContent(driver);
                    return LoginResult.Success;
                }

                waitingForOtc = GetOtcInput(driver) != null;

                if (!waitingForOtc)
                    throw new Exception($"Login page failed. {Reference.Login.UserId} not found.");
            }

            if (!waitingForOtc)
            {
                ThinkTime(1000);

                if (driver.IsVisible(By.Id("aadTile")))
                {
                    driver.FindElement(By.Id("aadTile")).Click(true);
                }

                ThinkTime(1000);

                //If expecting redirect then wait for redirect to trigger
                if (redirectAction != null)
                {
                    //Wait for redirect to occur.
                    Thread.Sleep(3000);

                    redirectAction.Invoke(new LoginRedirectEventArgs(username, password, driver));
                    return LoginResult.Redirect;
                }

                EnterPassword(driver, password);
                ThinkTime(1000);
            }

            EnterOneTimeCode(driver, mfaSecrectKey);

            if (mfaSecrectKey == null)
                ClickStaySignedIn(driver);

            ThinkTime(1000);

            var xpathToMainPage = By.XPath(Elements.Xpath[Reference.Login.CrmMainPage]);
            driver.WaitUntilVisible(xpathToMainPage
                , new TimeSpan(0, 0, 60),
                e => SwitchToDefaultContent(driver),
                f => throw new Exception($"Login page failed. {Reference.Login.CrmMainPage} not found."));

            return LoginResult.Success;
        }

        private static bool IsUserAlreadyLogged(IWebDriver driver)
        {
            var xpathToMainPage = By.XPath(Elements.Xpath[Reference.Login.CrmMainPage]);
            bool result = driver.HasElement(xpathToMainPage);
            return result;
        }

        private static string GenerateOneTimeCode(SecureString mfaSecrectKey)
        {
            // credits:
            // https://dev.to/j_sakamoto/selenium-testing---how-to-sign-in-to-two-factor-authentication-2joi
            // https://www.nuget.org/packages/Otp.NET/
            string key = mfaSecrectKey?.ToUnsecureString(); // <- this 2FA secret key.

            byte[] base32Bytes = Base32Encoding.ToBytes(key);

            var totp = new Totp(base32Bytes);
            var result = totp.ComputeTotp(); // <- got 2FA coed at this time!
            return result;
        }

        private bool EnterUserName(IWebDriver driver, SecureString username)
        {
            var input = driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Login.UserId]), new TimeSpan(0, 0, 30));
            if (input == null)
                return false;

            input.SendKeys(username.ToUnsecureString());
            input.SendKeys(Keys.Enter);
            return true;
        }

        private static void EnterPassword(IWebDriver driver, SecureString password)
        {
            var input = driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.LoginPassword]));
            input.SendKeys(password.ToUnsecureString());
            input.Submit();
        }

        private static void EnterOneTimeCode(IWebDriver driver, SecureString mfaSecrectKey)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    IWebElement input = GetOtcInput(driver);
                    var oneTimeCode = GenerateOneTimeCode(mfaSecrectKey);
                    input.SendKeys(oneTimeCode);
                    input.Submit();
                    return;
                }
                catch (Exception e)
                {
                    Trace.TraceWarning($"An Error ocur entering OTC. Exception: {e}");
                    if (attempts >= Constants.DefaultRetryAttempts)
                        throw;
                    Thread.Sleep(Constants.DefaultRetryDelay);
                    attempts++;
                }
            }
        }

        private static IWebElement GetOtcInput(IWebDriver driver)
            => driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.OneTimeCode]));

        private static void ClickStaySignedIn(IWebDriver driver)
        {
            driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn]), new TimeSpan(0, 0, 5));

            if (driver.IsVisible(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn])))
            {
                driver.ClickWhenAvailable(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn]));
            }
        }

        private static void SwitchToDefaultContent(IWebDriver driver)
        {
            driver.WaitForPageToLoad();
            driver.SwitchTo().Frame(0);
            driver.WaitForPageToLoad();

            //Switch Back to Default Content for Navigation Steps
            driver.SwitchTo().DefaultContent();
        }

        internal BrowserCommandResult<LoginResult> PassThroughLogin(Uri uri)
        {
            return Execute(GetOptions("Pass Through Login"), driver =>
            {
                driver.Navigate().GoToUrl(uri);

                driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Login.CrmMainPage])
                         , new TimeSpan(0, 0, 60),
                         e => SwitchToDefaultContent(driver),
                         f => throw new Exception("Login page failed."));

                return LoginResult.Success;
            });
        }

        public void ADFSLoginAction(LoginRedirectEventArgs args)

        {
            //Login Page details go here.  You will need to find out the id of the password field on the form as well as the submit button. 
            //You will also need to add a reference to the Selenium Webdriver to use the base driver. 
            //Example

            var d = args.Driver;

            d.FindElement(By.Id("passwordInput")).SendKeys(args.Password.ToUnsecureString());
            d.ClickWhenAvailable(By.Id("submitButton"), new TimeSpan(0, 0, 2));

            //Insert any additional code as required for the SSO scenario

            //Wait for CRM Page to load
            d.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Login.CrmMainPage])
                , new TimeSpan(0, 0, 60),
            e =>
            {
                d.WaitForPageToLoad();
                d.SwitchTo().Frame(0);
                d.WaitForPageToLoad();
            },
                f => throw new Exception("Login page failed."));

        }

        public void MSFTLoginAction(LoginRedirectEventArgs args)
        {
            //Login Page details go here.  You will need to find out the id of the password field on the form as well as the submit button. 
            //You will also need to add a reference to the Selenium Webdriver to use the base driver. 
            //Example

            var d = args.Driver;

            //d.FindElement(By.Id("passwordInput")).SendKeys(args.Password.ToUnsecureString());
            //d.ClickWhenAvailable(By.Id("submitButton"), new TimeSpan(0, 0, 2));

            //This method expects single sign-on

            Browser.ThinkTime(5000);

            d.WaitUntilVisible(By.XPath("//div[@id=\"mfaGreetingDescription\"]"));

            var AzureMFA = d.FindElement(By.XPath("//a[@id=\"WindowsAzureMultiFactorAuthentication\"]"));
            AzureMFA.Click(true);

            Thread.Sleep(20000);

            //Insert any additional code as required for the SSO scenario

            //Wait for CRM Page to load
            d.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Login.CrmMainPage])
                , new TimeSpan(0, 0, 60),
                e =>
                {
                    d.WaitForPageToLoad();
                    d.SwitchTo().Frame(0);
                    d.WaitForPageToLoad();
                },
                f => throw new Exception("Login page failed."));

        }

        #endregion

        #region Navigation
        internal BrowserCommandResult<bool> OpenApp(string appName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open App"), driver =>
            {
                driver.SwitchTo().DefaultContent();
                // Handle left hand Nav
                var xpathToAppMenu = By.XPath(AppElements.Xpath[AppReference.Navigation.AppMenuButton]);
                if (driver.HasElement(xpathToAppMenu))
                {
                    driver.ClickWhenAvailable(xpathToAppMenu);

                    var container = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.AppMenuContainer]));

                    var buttons = container.FindElements(By.TagName("button"));

                    var button = buttons.FirstOrDefault(x => x.Text.Trim() == appName);

                    if (button != null)
                        button.Click(true);
                    else
                        throw new InvalidOperationException($"App Name {appName} not found.");

                    driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Application.Shell]));
                    driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton]));
                    driver.WaitForPageToLoad();

                    driver.WaitForTransaction();

                    return true;
                }

                //Handle main.aspx?forceUCI=1
                // TODO: REMOVE COMMENTS FOR OLDER CRM Versions current v9.1
                //bool success = TryToClickInAppTile(driver, appName);
                //if (!success)
                //{
                //Switch to frame 0
                driver.SwitchTo().Frame(0);

                bool success = TryToClickInAppTile(driver, appName);
                if (!success)
                    throw new InvalidOperationException($"App Name {appName} not found.");
                //}

                InitializeModes();
                return true;
            });
        }

        private bool TryToClickInAppTile(IWebDriver driver, string appName)
        {
            var xpathToAppContainer = By.XPath(AppElements.Xpath[AppReference.Navigation.UCIAppContainer]);
            var xpathToAppTile = By.XPath(AppElements.Xpath[AppReference.Navigation.UCIAppTile].Replace("[NAME]", appName));

            var tileContainer = driver.WaitUntilAvailable(xpathToAppContainer, new TimeSpan(0, 0, 15));

            var appTile = tileContainer?.FindElement(xpathToAppTile);
            if (appTile == null)
                return false;

            appTile.Click(true);

            driver.WaitForTransaction();
            InitializeModes();
            return true;
        }

        internal BrowserCommandResult<bool> OpenGroupSubArea(string group, string subarea, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open Group Sub Area"), driver =>
            {
                //Make sure the sitemap-launcher is expanded - 9.1
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])))
                {
                    var expanded = bool.Parse(driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])).GetAttribute("aria-expanded"));

                    if (!expanded)
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton]));
                }

                var groups = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapMenuGroup]));
                var groupList = groups.FirstOrDefault(g => g.GetAttribute("aria-label").ToLowerString() == group.ToLowerString());
                if (groupList == null)
                {
                    throw new NotFoundException($"No group with the name '{group}' exists");
                }

                var subAreaItems = groupList.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapMenuItems]));
                var subAreaItem = subAreaItems.FirstOrDefault(a => a.GetAttribute("data-text").ToLowerString() == subarea.ToLowerString());
                if (subAreaItem == null)
                {
                    throw new NotFoundException($"No subarea with the name '{subarea}' exists inside of '{group}'");
                }

                subAreaItem.Click(true);

                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Grid.Container]));
                driver.WaitForPageToLoad();

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> OpenSubArea(string area, string subarea, int thinkTime = Constants.DefaultThinkTime)
        {
            //this.Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open Sub Area"), driver =>
            {
                area = area.ToLowerString();
                subarea = subarea.ToLowerString();

                //If the subarea is already in the left hand nav, click it
                var navSubAreas = OpenSubMenu(subarea, 100).Value;

                if (navSubAreas.ContainsKey(subarea))
                {
                    navSubAreas[subarea].Click(true);
                    driver.WaitForTransaction();

                    return true;
                }

                //We didn't find the subarea in the left hand nav. Try to find it
                var areas = OpenAreas(area).Value;

                IWebElement menuItem = null;
                bool foundMenuItem = areas.TryGetValue(area, out menuItem);

                if (foundMenuItem)
                    menuItem.Click(true);

                driver.WaitForTransaction();

                var subAreas = OpenSubMenu(subarea).Value;

                if (!subAreas.ContainsKey(subarea))
                    throw new InvalidOperationException($"No subarea with the name '{subarea}' exists inside of '{area}'.");

                subAreas[subarea].Click(true);

                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Grid.Container]));
                driver.WaitForPageToLoad();

                driver.WaitForTransaction();

                return true;
            });
        }

        public BrowserCommandResult<Dictionary<string, IWebElement>> OpenAreas(string area, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions("Open Unified Interface Area"), driver =>
            {
                //  9.1 ?? 9.0.2 <- inverted order (fallback first) run quickly
                var areas = OpenMenuFallback(area).Value ?? OpenMenu().Value;

                if (!areas.ContainsKey(area))
                    throw new InvalidOperationException($"No area with the name '{area}' exists.");

                return areas;
            });
        }
        public BrowserCommandResult<Dictionary<string, IWebElement>> OpenMenu(int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions("Open Menu"), driver =>
            {
                var dictionary = new Dictionary<string, IWebElement>();

                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.AreaButton]));

                driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.AreaMenu]),
                                            new TimeSpan(0, 0, 2),
                                            d =>
                                            {
                                                var menu = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.AreaMenu]));
                                                var menuItems = menu.FindElements(By.TagName("li"));
                                                foreach (var item in menuItems)
                                                {
                                                    dictionary.Add(item.Text.ToLowerString(), item);
                                                }
                                            },
                                            e => throw new InvalidOperationException("The Main Menu is not available."));


                return dictionary;
            });
        }
        public BrowserCommandResult<Dictionary<string, IWebElement>> OpenMenuFallback(string area, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions("Open Menu"), driver =>
            {
                var dictionary = new Dictionary<string, IWebElement>();

                //Make sure the sitemap-launcher is expanded - 9.1
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])))
                {
                    bool expanded = bool.Parse(driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton])).GetAttribute("aria-expanded"));

                    if (!expanded)
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton]));
                }

                //Is this the sitemap with enableunifiedinterfaceshellrefresh?
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapSwitcherButton])))
                {
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapSwitcherButton])).Click(true);

                    driver.WaitForTransaction();

                    driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapSwitcherFlyout]),
                        new TimeSpan(0, 0, 2),
                        d =>
                        {
                            var menu = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapSwitcherFlyout]));

                            var menuItems = menu.FindElements(By.TagName("li"));
                            foreach (var item in menuItems)
                            {
                                dictionary.Add(item.Text.ToLowerString(), item);
                            }
                        },
                        e => throw new InvalidOperationException("The Main Menu is not available."));
                }

                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapAreaMoreButton])))
                {
                    bool isVisible = driver.IsVisible(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapAreaMoreButton]));

                    if (isVisible)
                    {
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapAreaMoreButton]));

                        driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.AreaMoreMenu]),
                                               new TimeSpan(0, 0, 2),
                                               d =>
                                               {
                                                   var menu = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.AreaMoreMenu]));
                                                   var menuItems = menu.FindElements(By.TagName("li"));
                                                   foreach (var item in menuItems)
                                                   {
                                                       dictionary.Add(item.Text.ToLowerString(), item);
                                                   }
                                               },
                                               e =>
                                               {
                                                   throw new InvalidOperationException("The Main Menu is not available.");
                                               });
                    }
                    else
                    {
                        var singleItem = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapSingleArea].Replace("[NAME]", area)));

                        char[] trimCharacters = { '', '\r', '\n', '', '', '' };

                        dictionary.Add(singleItem.Text.Trim(trimCharacters).ToLowerString(), singleItem);

                    }
                }

                return dictionary;
            });
        }
        internal BrowserCommandResult<Dictionary<string, IWebElement>> OpenSubMenu(string subarea, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Open Sub Menu: {subarea}"), driver =>
            {
                var dictionary = new Dictionary<string, IWebElement>();

                //Sitemap without enableunifiedinterfaceshellrefresh
                if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Navigation.PinnedSitemapEntity])))
                {
                    bool isSiteMapLauncherCloseButtonVisible = driver.IsVisible(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherCloseButton]));

                    if (isSiteMapLauncherCloseButtonVisible)
                    {
                        // Close SiteMap launcher since it is open
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherCloseButton]));
                    }

                    driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SiteMapLauncherButton]));

                    var menuContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SubAreaContainer]));

                    var subItems = menuContainer.FindElements(By.TagName("li"));

                    foreach (var subItem in subItems)
                    {
                        // Check 'Id' attribute, NULL value == Group Header
                        if (!string.IsNullOrEmpty(subItem.GetAttribute("Id")))
                        {
                            // Filter out duplicate entity keys - click the first one in the list
                            if (!dictionary.ContainsKey(subItem.Text.ToLowerString()))
                                dictionary.Add(subItem.Text.ToLowerString(), subItem);
                        }
                    }
                }

                else
                {
                    //Sitemap with enableunifiedinterfaceshellrefresh enabled
                    var menuShell = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.SubAreaContainer]));

                    //The menu is broke into multiple sections. Gather all items.
                    foreach (IWebElement menuSection in menuShell)
                    {
                        var menuItems = menuSection.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.SitemapMenuItems]));

                        foreach (var menuItem in menuItems)
                        {
                            if (!string.IsNullOrEmpty(menuItem.Text))
                            {
                                if (!dictionary.ContainsKey(menuItem.Text.ToLower()))
                                    dictionary.Add(menuItem.Text.ToLower(), menuItem);
                            }
                        }
                    }
                }

                return dictionary;
            });
        }

        internal BrowserCommandResult<bool> OpenSettingsOption(string command, string dataId, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Open " + command + " " + dataId), driver =>
            {
                var cmdButtonBar = AppElements.Xpath[AppReference.Navigation.SettingsLauncherBar].Replace("[NAME]", command);
                var cmdLauncher = AppElements.Xpath[AppReference.Navigation.SettingsLauncher].Replace("[NAME]", command);

                if (!driver.IsVisible(By.XPath(cmdLauncher)))
                {
                    driver.ClickWhenAvailable(By.XPath(cmdButtonBar));

                    Thread.Sleep(1000);

                    driver.SetVisible(By.XPath(cmdLauncher), true);
                    driver.WaitUntilVisible(By.XPath(cmdLauncher));
                }

                var menuContainer = driver.FindElement(By.XPath(cmdLauncher));
                var menuItems = menuContainer.FindElements(By.TagName("button"));
                var button = menuItems.FirstOrDefault(x => x.GetAttribute("data-id").Contains(dataId));

                if (button != null)
                {
                    button.Click();
                }
                else
                {
                    throw new InvalidOperationException($"No command with the exists inside of the Command Bar.");
                }

                return true;
            });
        }

        /// <summary>
        /// Opens the Guided Help
        /// </summary>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.Navigation.OpenGuidedHelp();</example>
        public BrowserCommandResult<bool> OpenGuidedHelp(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Open Guided Help"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.GuidedHelp]));

                return true;
            });
        }

        /// <summary>
        /// Opens the Admin Portal
        /// </summary>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.Navigation.OpenAdminPortal();</example>
        internal BrowserCommandResult<bool> OpenAdminPortal(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);
            return Execute(GetOptions("Open Admin Portal"), driver =>
            {
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Application.Shell]));
                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.AdminPortal]))?.Click();
                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.AdminPortalButton]))?.Click();
                return true;
            });
        }

        /// <summary>
        /// Open Global Search
        /// </summary>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.Navigation.OpenGlobalSearch();</example>
        internal BrowserCommandResult<bool> OpenGlobalSearch(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Open Global Search"), driver =>
            {
                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Navigation.SearchButton]),
                new TimeSpan(0, 0, 5),
                d => { driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Navigation.SearchButton])); },
                d => throw new InvalidOperationException("The Global Search button is not available."));
                return true;
            });
        }
        internal BrowserCommandResult<bool> ClickQuickLaunchButton(string toolTip, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Quick Launch: {toolTip}"), driver =>
            {
                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickLaunchMenu]));

                //Text could be in the crumb bar.  Find the Quick launch bar buttons and click that one.
                var buttons = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickLaunchMenu]));
                var launchButton = buttons.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickLaunchButton].Replace("[NAME]", toolTip)));
                launchButton.Click();

                return true;
            });
        }

        internal BrowserCommandResult<bool> QuickCreate(string entityName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Quick Create: {entityName}"), driver =>
            {
                //Click the + button in the ribbon
                var quickCreateButton = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickCreateButton]));
                quickCreateButton.Click(true);

                //Find the entity name in the list
                var entityMenuList = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickCreateMenuList]));
                var entityMenuItems = entityMenuList.FindElements(By.XPath(AppElements.Xpath[AppReference.Navigation.QuickCreateMenuItems]));
                var entitybutton = entityMenuItems.FirstOrDefault(e => e.Text.Contains(entityName, StringComparison.OrdinalIgnoreCase));

                if (entitybutton == null)
                    throw new Exception(string.Format("{0} not found in Quick Create list.", entityName));

                //Click the entity name
                entitybutton.Click(true);

                driver.WaitForTransaction();

                return true;
            });
        }

        #endregion

        #region Dialogs
        internal bool SwitchToDialog(int frameIndex = 0)
        {
            var index = "";
            if (frameIndex > 0)
                index = frameIndex.ToString();

            Browser.Driver.SwitchTo().DefaultContent();

            // Check to see if dialog is InlineDialog or popup
            var inlineDialog = Browser.Driver.HasElement(By.XPath(Elements.Xpath[Reference.Frames.DialogFrame].Replace("[INDEX]", index)));
            if (inlineDialog)
            {
                //wait for the content panel to render
                Browser.Driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Frames.DialogFrame].Replace("[INDEX]", index)),
                                                  new TimeSpan(0, 0, 2),
                                                  d => { Browser.Driver.SwitchTo().Frame(Elements.ElementId[Reference.Frames.DialogFrameId].Replace("[INDEX]", index)); });
                return true;
            }
            else
            {
                // need to add this functionality
                //SwitchToPopup();
            }
            return true;
        }
        internal BrowserCommandResult<bool> CloseWarningDialog()
        {
            return Execute(GetOptions($"Close Warning Dialog"), driver =>
            {
                var inlineDialog = SwitchToDialog();
                if (inlineDialog)
                {
                    var dialogFooter = driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Dialogs.WarningFooter]));

                    if (
                        !(dialogFooter?.FindElements(By.XPath(Elements.Xpath[Reference.Dialogs.WarningCloseButton])).Count >
                          0)) return true;
                    var closeBtn = dialogFooter.FindElement(By.XPath(Elements.Xpath[Reference.Dialogs.WarningCloseButton]));
                    closeBtn.Click();
                }
                return true;
            });
        }
        internal BrowserCommandResult<bool> ConfirmationDialog(bool ClickConfirmButton)
        {
            //Passing true clicks the confirm button.  Passing false clicks the Cancel button.
            return Execute(GetOptions($"Confirm or Cancel Confirmation Dialog"), driver =>
            {
                var inlineDialog = SwitchToDialog();
                if (inlineDialog)
                {
                    //Wait until the buttons are available to click
                    var dialogFooter = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Dialogs.ConfirmButton]));

                    if (
                        !(dialogFooter?.FindElements(By.XPath(AppElements.Xpath[AppReference.Dialogs.ConfirmButton])).Count >
                          0)) return true;

                    //Click the Confirm or Cancel button
                    IWebElement buttonToClick;
                    if (ClickConfirmButton)
                        buttonToClick = dialogFooter.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.ConfirmButton]));
                    else
                        buttonToClick = dialogFooter.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.CancelButton]));

                    buttonToClick.Click();
                }
                return true;
            });
        }
        internal BrowserCommandResult<bool> AssignDialog(Dialogs.AssignTo to, string userOrTeamName)
        {
            return Execute(GetOptions($"Assign to User or Team Dialog"), driver =>
            {
                var inlineDialog = SwitchToDialog();
                if (inlineDialog)
                {
                    if (to != Dialogs.AssignTo.Me)
                    {
                        //Click the Option to Assign to User Or Team
                        driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Dialogs.AssignDialogToggle]));

                        var toggleButton = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Dialogs.AssignDialogToggle]), "Me/UserTeam toggle button unavailable");
                        if (toggleButton.Text == "Me")
                            toggleButton.Click();

                        //Set the User Or Team
                        var userOrTeamField = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookup]), "User field unavailable");

                        if (userOrTeamField.FindElements(By.TagName("input")).Count > 0)
                        {
                            var input = userOrTeamField.FindElement(By.TagName("input"));
                            if (input != null)
                            {
                                input.Click();

                                driver.WaitForTransaction();

                                input.SendKeys(userOrTeamName, true);
                            }
                        }

                        //Pick the User from the list
                        driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Dialogs.AssignDialogUserTeamLookupResults]));

                        driver.WaitForTransaction();

                        var container = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.AssignDialogUserTeamLookupResults]));
                        var records = container.FindElements(By.TagName("li"));
                        foreach (var record in records)
                        {
                            if (record.Text.StartsWith(userOrTeamName, StringComparison.OrdinalIgnoreCase))
                                record.Click(true);
                        }
                    }

                    //Click Assign
                    var okButton = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.AssignDialogOKButton]));
                    okButton.Click(true);

                }
                return true;
            });
        }
        internal BrowserCommandResult<bool> SwitchProcessDialog(string processToSwitchTo)
        {
            return Execute(GetOptions($"Switch Process Dialog"), driver =>
            {
                //Wait for the Grid to load
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Dialogs.ActiveProcessGridControlContainer]));

                //Select the Process
                var popup = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.SwitchProcessContainer]));
                var labels = popup.FindElements(By.TagName("label"));
                foreach (var label in labels)
                {
                    if (label.Text.Equals(processToSwitchTo, StringComparison.OrdinalIgnoreCase))
                    {
                        label.Click();
                        break;
                    }
                }

                //Click the OK button
                var okBtn = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Dialogs.SwitchProcessDialogOK]));
                okBtn.Click();

                return true;
            });
        }
        internal BrowserCommandResult<bool> CloseOpportunityDialog(bool clickOK)
        {
            return Execute(GetOptions($"Close Opportunity Dialog"), driver =>
            {
                var inlineDialog = SwitchToDialog();

                if (inlineDialog)
                {
                    //Close Opportunity
                    var xPath = AppElements.Xpath[AppReference.Dialogs.CloseOpportunity.Ok];

                    //Cancel
                    if (!clickOK)
                        xPath = AppElements.Xpath[AppReference.Dialogs.CloseOpportunity.Ok];

                    driver.WaitUntilClickable(By.XPath(xPath),
                        new TimeSpan(0, 0, 5),
                        d => { driver.ClickWhenAvailable(By.XPath(xPath)); },
                        d => { throw new InvalidOperationException("The Close Opportunity dialog is not available."); });
                }
                return true;
            });
        }
        internal BrowserCommandResult<bool> HandleSaveDialog()
        {
            //If you click save and something happens, handle it.  Duplicate Detection/Errors/etc...
            //Check for Dialog and figure out which type it is and return the dialog type.

            //Introduce think time to avoid timing issues on save dialog
            Browser.ThinkTime(1000);

            return Execute(GetOptions($"Validate Save"), driver =>
            {
                //Is it Duplicate Detection?
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.DuplicateDetectionWindowMarker])))
                {
                    if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.DuplicateDetectionGridRows])))
                    {
                        //Select the first record in the grid
                        driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.DuplicateDetectionGridRows]))[0].Click(true);

                        //Click Ignore and Save
                        driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.DuplicateDetectionIgnoreAndSaveButton])).Click(true);
                        driver.WaitForTransaction();
                    }
                }

                //Is it an Error?
                if (driver.HasElement(By.XPath("//div[contains(@data-id,'errorDialogdialog')]")))
                {
                    var errorDialog = driver.FindElement(By.XPath("//div[contains(@data-id,'errorDialogdialog')]"));

                    var errorDetails = errorDialog.FindElement(By.XPath(".//*[contains(@data-id,'errorDialog_subtitle')]"));

                    if (!string.IsNullOrEmpty(errorDetails.Text))
                        throw new InvalidOperationException(errorDetails.Text);
                }


                return true;
            });
        }

        /// <summary>
        /// Opens the dialog
        /// </summary>
        /// <param name="dialog"></param>
        public BrowserCommandResult<List<ListItem>> LookupResultsDropdown(IWebElement dialog)
        {
            var list = new List<ListItem>();
            var dialogItems = dialog.FindElements(By.XPath(".//li"));

            foreach (var dialogItem in dialogItems)
            {
                var titleLinks = dialogItem.FindElements(By.XPath(".//label"));
                var divLinks = dialogItem.FindElements(By.XPath(".//div"));

                if (titleLinks != null && titleLinks.Count > 0 && divLinks != null && divLinks.Count > 0)
                {
                    var title = titleLinks[0].GetAttribute("innerText");
                    var divId = divLinks[0].GetAttribute("id");

                    list.Add(new ListItem()
                    {
                        Title = title,
                        Id = divId,
                    });
                }
            }

            return list;
        }

        /// <summary>
        /// Opens the dialog
        /// </summary>
        /// <param name="dialog"></param>
        public BrowserCommandResult<List<ListItem>> OpenDialog(IWebElement dialog)
        {
            var list = new List<ListItem>();
            var dialogItems = dialog.FindElements(By.XPath(".//li"));

            foreach (var dialogItem in dialogItems)
            {
                var titleLinks = dialogItem.FindElements(By.XPath(".//label/span"));
                var divLinks = dialogItem.FindElements(By.XPath(".//div"));

                if (titleLinks != null && titleLinks.Count > 0 && divLinks != null && divLinks.Count > 0)
                {
                    var title = titleLinks[0].GetAttribute("innerText");
                    var divId = divLinks[0].GetAttribute("id");

                    list.Add(new ListItem()
                    {
                        Title = title,
                        Id = divId,
                    });
                }
            }

            return list;
        }
        #endregion

        #region CommandBar
        internal BrowserCommandResult<bool> ClickCommand(string name, string subname = "", bool moreCommands = false, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Click Command"), driver =>
            {
                IWebElement ribbon = null;

                //Find the button in the CommandBar
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Container])))
                    ribbon = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Container]));

                if (ribbon == null)
                {
                    if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.ContainerGrid])))
                        ribbon = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.ContainerGrid]));
                    else
                        throw new InvalidOperationException("Unable to find the ribbon.");
                }

                //Get the CommandBar buttons
                var items = ribbon.FindElements(By.TagName("li"));

                //Is the button in the ribbon?
                if (items.Any(x => x.GetAttribute("aria-label").Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    items.FirstOrDefault(x => x.GetAttribute("aria-label").Equals(name, StringComparison.OrdinalIgnoreCase)).Click(true);
                    driver.WaitForTransaction();
                }
                else
                {
                    //Is the button in More Commands?
                    if (items.Any(x => x.GetAttribute("aria-label").Equals("More Commands", StringComparison.OrdinalIgnoreCase)))
                    {
                        //Click More Commands
                        items.FirstOrDefault(x => x.GetAttribute("aria-label").Equals("More Commands", StringComparison.OrdinalIgnoreCase)).Click(true);
                        driver.WaitForTransaction();

                        //Click the button
                        if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Button].Replace("[NAME]", name))))
                        {
                            driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Button].Replace("[NAME]", name))).Click(true);
                            driver.WaitForTransaction();
                        }
                        else
                            throw new InvalidOperationException($"No command with the name '{name}' exists inside of Commandbar.");
                    }
                    else
                        throw new InvalidOperationException($"No command with the name '{name}' exists inside of Commandbar.");
                }

                if (!string.IsNullOrEmpty(subname))
                {
                    var submenu = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.CommandBar.MoreCommandsMenu]));

                    var subbutton = submenu.FindElements(By.TagName("button")).FirstOrDefault(x => x.Text == subname);

                    if (subbutton != null)
                    {
                        subbutton.Click(true);
                    }
                    else
                        throw new InvalidOperationException($"No sub command with the name '{subname}' exists inside of Commandbar.");

                }

                driver.WaitForTransaction();

                return true;
            });
        }

        /// <summary>
        /// Returns the values of CommandBar objects
        /// </summary>
        /// <param name="moreCommands">The moreCommands</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.Related.ClickCommand("ADD NEW CASE");</example>
        internal BrowserCommandResult<List<string>> GetCommandValues(bool includeMoreCommandsValues = false, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Get CommandBar Command Count"), driver =>
            {
                IWebElement ribbon = null;
                List<string> commandValues = new List<string>();

                //Find the button in the CommandBar
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Container])))
                    ribbon = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.Container]));

                if (ribbon == null)
                {
                    if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.ContainerGrid])))
                        ribbon = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.ContainerGrid]));
                    else
                        throw new InvalidOperationException("Unable to find the ribbon.");
                }

                //Get the CommandBar buttons
                var commandBarItems = ribbon.FindElements(By.TagName("li"));

                foreach (var value in commandBarItems)
                {
                    if (value.Text != "")
                    {
                        string commandText = value.Text.ToString();

                        if (commandText.Contains("\r\n"))
                        {
                            commandText = commandText.Substring(0, commandText.IndexOf("\r\n"));
                        }

                        if (!commandValues.Contains(value.Text))
                        {
                            commandValues.Add(commandText);
                        }
                    }
                }

                if (includeMoreCommandsValues)
                {
                    if (commandBarItems.Any(x => x.GetAttribute("aria-label").Equals("More Commands", StringComparison.OrdinalIgnoreCase)))
                    {
                        //Click More Commands Button
                        commandBarItems.FirstOrDefault(x => x.GetAttribute("aria-label").Equals("More Commands", StringComparison.OrdinalIgnoreCase)).Click(true);
                        driver.WaitForTransaction();

                        //Click the button
                        var moreCommandsMenu = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.CommandBar.MoreCommandsMenu]));

                        if (moreCommandsMenu != null)
                        {
                            var moreCommandsItems = moreCommandsMenu.FindElements(By.TagName("li"));

                            foreach (var value in moreCommandsItems)
                            {
                                if (value.Text != "")
                                {
                                    string commandText = value.Text.ToString();

                                    if (commandText.Contains("\r\n"))
                                    {
                                        commandText = commandText.Substring(0, commandText.IndexOf("\r\n"));
                                    }

                                    if (!commandValues.Contains(value.Text))
                                    {
                                        commandValues.Add(commandText);
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw new InvalidOperationException("Unable to locate the 'More Commands' menu");
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException("No button matching 'More Commands' exists in the CommandBar");
                    }
                }

                return commandValues;
            });
        }

        #endregion

        #region Grid
        public BrowserCommandResult<Dictionary<string, string>> OpenViewPicker(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open View Picker"), driver =>
            {
                var dictionary = new Dictionary<string, string>();

                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Grid.ViewSelector]),
                    new TimeSpan(0, 0, 20),
                    e => e.Click(),
                    d => throw new Exception("Unable to click the View Picker"));

                var viewContainer = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.ViewContainer]));
                var viewItems = viewContainer.FindElements(By.TagName("li"));

                foreach (var viewItem in viewItems)
                {
                    if (viewItem.GetAttribute("role") != null && viewItem.GetAttribute("role") == "option")
                    {
                        dictionary.Add(viewItem.Text, viewItem.GetAttribute("id"));
                    }
                }

                return dictionary;
            });
        }
        internal BrowserCommandResult<bool> SwitchView(string viewName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Switch View"), driver =>
            {
                var views = OpenViewPicker().Value;
                Thread.Sleep(500);
                if (!views.ContainsKey(viewName))
                {
                    throw new InvalidOperationException($"No view with the name '{viewName}' exists.");
                }

                var viewId = views[viewName];
                driver.ClickWhenAvailable(By.Id(viewId));

                return true;
            });
        }
        internal BrowserCommandResult<bool> OpenRecord(int index, int thinkTime = Constants.DefaultThinkTime, bool checkRecord = false)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Open Grid Record"), driver =>
            {
                var currentindex = 0;
                //var control = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.Container]));

                var rows = driver.FindElements(By.ClassName("wj-row"));

                //TODO: The grid only has a small subset of records. Need to load them all
                foreach (var row in rows)
                {
                    if (!string.IsNullOrEmpty(row.GetAttribute("data-lp-id")))
                    {
                        if (currentindex == index)
                        {
                            var tag = checkRecord ? "div" : "a";
                            row.FindElement(By.TagName(tag)).Click();
                            break;
                        }

                        currentindex++;
                    }
                }

                //driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Entity.Form]));

                driver.WaitForTransaction();

                return true;
            });
        }
        internal BrowserCommandResult<bool> Search(string searchCriteria, bool clearByDefault = true, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Search"), driver =>
            {
                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind]));

                if (clearByDefault)
                {
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind])).Clear();
                }

                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind])).SendKeys(searchCriteria);
                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind])).SendKeys(Keys.Enter);

                //driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> ClearSearch(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Clear Search"), driver =>
            {
                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind]));

                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.QuickFind])).Clear();

                return true;
            });
        }

        internal BrowserCommandResult<List<GridItem>> GetGridItems(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Get Grid Items"), driver =>
            {
                var returnList = new List<GridItem>();

                driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.Container]));

                var rows = driver.FindElements(By.ClassName("wj-row"));
                var columnGroup = driver.FindElement(By.ClassName("wj-colheaders"));

                foreach (var row in rows)
                {
                    if (!string.IsNullOrEmpty(row.GetAttribute("data-lp-id")) && !string.IsNullOrEmpty(row.GetAttribute("role")))
                    {
                        //MscrmControls.Grid.ReadOnlyGrid|entity_control|account|00000000-0000-0000-00aa-000010001001|account|cc-grid|grid-cell-container
                        var datalpid = row.GetAttribute("data-lp-id").Split('|');
                        var cells = row.FindElements(By.ClassName("wj-cell"));
                        var currentindex = 0;
                        var link =
                           $"{new Uri(driver.Url).Scheme}://{new Uri(driver.Url).Authority}/main.aspx?etn={datalpid[2]}&pagetype=entityrecord&id=%7B{datalpid[3]}%7D";

                        var item = new GridItem
                        {
                            EntityName = datalpid[2],
                            Url = new Uri(link)
                        };

                        foreach (var column in columnGroup.FindElements(By.ClassName("wj-row")))
                        {
                            var rowHeaders = column.FindElements(By.TagName("div"))
                                .Where(c => !string.IsNullOrEmpty(c.GetAttribute("title")) && !string.IsNullOrEmpty(c.GetAttribute("id")));

                            foreach (var header in rowHeaders)
                            {
                                var id = header.GetAttribute("id");
                                var className = header.GetAttribute("class");
                                var cellData = cells[currentindex + 1].GetAttribute("title");

                                if (!string.IsNullOrEmpty(id)
                                    && className.Contains("wj-cell")
                                    && !string.IsNullOrEmpty(cellData)
                                    && cells.Count > currentindex
                                )
                                {
                                    item[id] = cellData.Replace("-", "");
                                }
                                currentindex++;
                            }
                            returnList.Add(item);
                        }
                    }
                }
                return returnList;
            });
        }
        internal BrowserCommandResult<bool> NextPage(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Next Page"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.NextPage]));

                driver.WaitForTransaction();

                return true;
            });
        }
        internal BrowserCommandResult<bool> PreviousPage(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Previous Page"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.PreviousPage]));

                driver.WaitForTransaction();

                return true;
            });
        }
        internal BrowserCommandResult<bool> FirstPage(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"First Page"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.FirstPage]));

                driver.WaitForTransaction();

                return true;
            });
        }
        internal BrowserCommandResult<bool> SelectAll(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Select All"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.SelectAll]));

                driver.WaitForTransaction();

                return true;
            });
        }
        public BrowserCommandResult<bool> ShowChart(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Show Chart"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Grid.ShowChart])))
                {
                    driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.ShowChart]));

                    driver.WaitForTransaction();
                }
                else
                {
                    throw new Exception("The Show Chart button does not exist.");
                }

                return true;
            });
        }
        public BrowserCommandResult<bool> HideChart(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Hide Chart"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Grid.HideChart])))
                {
                    driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.HideChart]));

                    driver.WaitForTransaction();
                }
                else
                {
                    throw new Exception("The Hide Chart button does not exist.");
                }

                return true;
            });
        }
        public BrowserCommandResult<bool> FilterByLetter(char filter, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            if (!char.IsLetter(filter) && filter != '#')
                throw new InvalidOperationException("Filter criteria is not valid.");

            return Execute(GetOptions("Filter by Letter"), driver =>
            {
                var jumpBar = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.JumpBar]));
                var link = jumpBar.FindElement(By.Id(filter + "_link"));

                if (link != null)
                {
                    link.Click();

                    driver.WaitForTransaction();
                }
                else
                {
                    throw new Exception($"Filter with letter: {filter} link does not exist");
                }

                return true;
            });
        }
        public BrowserCommandResult<bool> FilterByAll(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Filter by All Records"), driver =>
            {
                var jumpBar = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.JumpBar]));
                var link = jumpBar.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.FilterByAll]));

                if (link != null)
                {
                    link.Click();

                    driver.WaitForTransaction();
                }
                else
                {
                    throw new Exception($"Filter by All link does not exist");
                }

                return true;
            });
        }
        public BrowserCommandResult<bool> SelectRecord(int index, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Select Grid Record"), driver =>
            {
                var container = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.RowsContainer]),
                                                        $"Grid Container does not exist.");

                var row = container.FindElement(By.Id("id-cell-" + index + "-1"));

                if (row != null)
                    row.Click();
                else
                    throw new Exception($"Row with index: {index} does not exist.");

                return false;
            });
        }
        public BrowserCommandResult<bool> SwitchChart(string chartName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            if (!Browser.Driver.IsVisible(By.XPath(AppElements.Xpath[AppReference.Grid.ChartSelector])))
                ShowChart();

            Browser.ThinkTime(1000);

            return Execute(GetOptions("Switch Chart"), driver =>
            {
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Grid.ChartSelector]));

                var list = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.ChartViewList]));

                driver.ClickWhenAvailable(By.XPath("//li[contains(@title,'" + chartName + "')]"));

                return true;
            });
        }
        public BrowserCommandResult<bool> Sort(string columnName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Sort by {columnName}"), driver =>
            {
                var sortCol = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.GridSortColumn].Replace("[COLNAME]", columnName)));

                if (sortCol == null)
                    throw new InvalidOperationException($"Column: {columnName} Does not exist");

                sortCol.Click(true);
                driver.WaitForTransaction();
                return true;
            });
        }
        #endregion

        #region RelatedGrid
        /// <summary>
        /// Opens the grid record.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        public BrowserCommandResult<bool> OpenGridRow(int index, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open Grid Item"), driver =>
            {
                var grid = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.Container]));
                var gridCellContainer = grid.FindElement(By.XPath(AppElements.Xpath[AppReference.Grid.CellContainer]));
                var rowCount = gridCellContainer.GetAttribute("data-row-count");
                var count = 0;

                if (rowCount == null || !int.TryParse(rowCount, out count) || count <= 0) return true;
                var link =
                    gridCellContainer.FindElement(
                        By.XPath("//div[@role='gridcell'][@header-row-number='" + index + "']/following::div"));

                if (link != null)
                {
                    link.Click();
                }
                else
                {
                    throw new InvalidOperationException($"No record with the index '{index}' exists.");
                }

                driver.WaitForTransaction();
                return true;
            });
        }

        public BrowserCommandResult<bool> ClickRelatedCommand(string name, string subName = null)
        {
            return Execute(GetOptions("Click Related Tab Command"), driver =>
            {
                if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Related.CommandBarButton].Replace("[NAME]", name))))
                    throw new NotFoundException($"{name} button not found. Button names are case sensitive. Please check for proper casing of button name.");

                driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Related.CommandBarButton].Replace("[NAME]", name))).Click(true);

                driver.WaitForTransaction();

                if (subName != null)
                {
                    if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Related.CommandBarSubButton].Replace("[NAME]", subName))))
                        throw new NotFoundException($"{subName} button not found");

                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Related.CommandBarSubButton].Replace("[NAME]", subName))).Click(true);

                    driver.WaitForTransaction();
                }
                return true;
            });
        }
        #endregion

        #region Entity

        internal BrowserCommandResult<bool> CancelQuickCreate(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Cancel Quick Create"), driver =>
            {
                var save = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.QuickCreate.CancelButton]),
                    "Quick Create Cancel Button is not available");
                save?.Click(true);

                driver.WaitForTransaction();

                return true;
            });
        }

        /// <summary>
        /// Open Entity
        /// </summary>
        /// <param name="entityName">The entity name</param>
        /// <param name="id">The Id</param>
        /// <param name="thinkTime">The think time</param>
        internal BrowserCommandResult<bool> OpenEntity(string entityName, Guid id, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Open: {entityName} {id}"), driver =>
            {
                //https:///main.aspx?appid=98d1cf55-fc47-e911-a97c-000d3ae05a70&pagetype=entityrecord&etn=lead&id=ed975ea3-531c-e511-80d8-3863bb3ce2c8
                var uri = new Uri(Browser.Driver.Url);
                var qs = HttpUtility.ParseQueryString(uri.Query.ToLower());
                var appId = qs.Get("appid");
                var link = $"{uri.Scheme}://{uri.Authority}/main.aspx?appid={appId}&etn={entityName}&pagetype=entityrecord&id={id}";

                if (Browser.Options.UCITestMode)
                {
                    link += "&flags=testmode=true";
                }

                driver.Navigate().GoToUrl(link);

                //SwitchToContent();
                driver.WaitForPageToLoad();
                driver.WaitForTransaction();
                driver.WaitUntilClickable(By.XPath(Elements.Xpath[Reference.Entity.Form]),
                    new TimeSpan(0, 0, 30),
                    null,
                    d => { throw new Exception("CRM Record is Unavailable or not finished loading. Timeout Exceeded"); }
                );

                return true;
            });
        }

        /// <summary>
        /// Saves the entity
        /// </summary>
        /// <param name="thinkTime"></param>
        internal BrowserCommandResult<bool> Save(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Save"), driver =>
            {
                var save = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.Save]),
                    "Save Buttton is not available");

                save?.Click();

                return true;
            });
        }

        internal BrowserCommandResult<bool> SaveQuickCreate(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"SaveQuickCreate"), driver =>
            {
                var save = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.QuickCreate.SaveAndCloseButton]),
                    "Quick Create Save Button is not available");
                save?.Click(true);

                driver.WaitForTransaction();

                return true;
            });
        }

        /// <summary>
        /// Open record set and navigate record index.
        /// This method supersedes Navigate Up and Navigate Down outside of UCI 
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.Entity.OpenRecordSetNavigator();</example>
        public BrowserCommandResult<bool> OpenRecordSetNavigator(int index = 0, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Open Record Set Navigator"), driver =>
            {
                // check if record set navigator parent div is set to open
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.RecordSetNavigatorOpen])))
                {
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.RecordSetNavigator])).Click();
                }

                var navList = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.RecordSetNavList]));
                var links = navList.FindElements(By.TagName("li"));
                try
                {
                    links[index].Click();
                }
                catch
                {
                    throw new InvalidOperationException($"No record with the index '{index}' exists.");
                }

                driver.WaitForPageToLoad();

                return true;
            });
        }

        /// <summary>
        /// Close Record Set Navigator
        /// </summary>
        /// <param name="thinkTime"></param>
        /// <example>xrmApp.Entity.CloseRecordSetNavigator();</example>
        public BrowserCommandResult<bool> CloseRecordSetNavigator(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions("Close Record Set Navigator"), driver =>
            {
                var closeSpan =
                    driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.RecordSetNavCollapseIcon]));

                if (closeSpan)
                {
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.RecordSetNavCollapseIconParent])).Click();
                }

                return true;
            });
        }

        /// <summary>
        /// Set Value
        /// </summary>
        /// <param name="field">The field</param>
        /// <param name="value">The value</param>
        /// <example>xrmApp.Entity.SetValue("firstname", "Test");</example>
        internal BrowserCommandResult<bool> SetValue(string field, string value)
        {
            return Execute(GetOptions($"Set Value"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", field)));

                if (fieldContainer.FindElements(By.TagName("input")).Count > 0)
                {
                    var input = fieldContainer.FindElement(By.TagName("input"));
                    if (input != null)
                    {
                        input.Click();

                        if (string.IsNullOrEmpty(value))
                        {
                            input.SendKeys(Keys.Control + "a");
                            input.SendKeys(Keys.Backspace);
                        }
                        else
                        {
                            input.SendKeys(value, true);
                        }
                    }
                }
                else if (fieldContainer.FindElements(By.TagName("textarea")).Count > 0)
                {
                    if (string.IsNullOrEmpty(value))
                    {
                        fieldContainer.FindElement(By.TagName("textarea")).SendKeys(Keys.Control + "a");
                        fieldContainer.FindElement(By.TagName("textarea")).SendKeys(Keys.Backspace);
                    }
                    else
                    {
                        fieldContainer.FindElement(By.TagName("textarea")).SendKeys(value, true);
                    }
                }
                else
                {
                    throw new Exception($"Field with name {field} does not exist.");
                }

                // Needed to transfer focus out of special fields (email or phone)
                driver.FindElement(By.TagName("body")).Click();

                return true;
            });
        }

        /// <summary>
        /// Sets the value of a Lookup, Customer, Owner or ActivityParty Lookup which accepts only a single value.
        /// </summary>
        /// <param name="control">The lookup field name, value or index of the lookup.</param>
        /// <example>xrmApp.Entity.SetValue(new Lookup { Name = "prrimarycontactid", Value = "Rene Valdes (sample)" });</example>
        /// The default index position is 0, which will be the first result record in the lookup results window. Suppy a value > 0 to select a different record if multiple are present.
        internal BrowserCommandResult<bool> SetValue(LookupItem control, int index = 0)
        {
            return Execute(GetOptions($"Set Lookup Value: {control.Name}"), driver =>
            {
                driver.WaitForTransaction(120);

                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupFieldContainer].Replace("[NAME]", control.Name)));


                ClearValue(control);

                var input = fieldContainer.FindElements(By.TagName("input")).Count > 0
                    ? fieldContainer.FindElement(By.TagName("input"))
                    : null;

                if (input != null)
                {
                    input.SendKeys(Keys.Control + "a");
                    input.SendKeys(Keys.Backspace);
                    input.SendKeys(control.Value, true);

                    //No longer needed, the search dialog opens when you enter the value
                    //var byXPath = By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupSearchButton].Replace("[NAME]", control.Name));
                    //driver.ClickWhenAvailable(byXPath);
                    driver.WaitForTransaction();
                }

                if (!string.IsNullOrEmpty(control.Value))
                {
                    SetLookUpByValue(driver, control, index);
                }
                else if (control.Value == "")
                {
                    SetLookupByIndex(driver, control, index);
                }
                else if (control.Value == null)
                {
                    throw new InvalidOperationException($"No value has been provided for the LookupItem {control.Name}. Please provide a value or an empty string and try again.");
                }

                return true;
            });
        }

        /// <summary>
        /// Sets the value of an ActivityParty Lookup.
        /// </summary>
        /// <param name="control">The lookup field name, value or index of the lookup.</param>
        /// <example>xrmApp.Entity.SetValue(new Lookup[] { Name = "to", Value = "Rene Valdes (sample)" }, { Name = "to", Value = "Alpine Ski House (sample)" } );</example>
        /// The default index position is 0, which will be the first result record in the lookup results window. Suppy a value > 0 to select a different record if multiple are present.
        internal BrowserCommandResult<bool> SetValue(LookupItem[] controls, int index = 0, bool clearFirst = true)
        {
            return Execute(GetOptions($"Set ActivityParty Lookup Value: {controls.First().Name}"), driver =>
            {
                driver.WaitForTransaction(5);

                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupFieldContainer].Replace("[NAME]", controls.First().Name)));

                if (clearFirst)
                    ClearValue(controls[0]);

                var input = fieldContainer.FindElements(By.TagName("input")).Count > 0
                    ? fieldContainer.FindElement(By.TagName("input"))
                    : null;

                foreach (var control in controls)
                {
                    if (input != null)
                    {
                        if (!string.IsNullOrEmpty(control.Value))
                        {
                            input.SendKeys(control.Value, true);
                            input.SendKeys(Keys.Tab);
                            input.SendKeys(Keys.Enter);
                        }
                        else
                        {
                            input.Click();
                        }
                        driver.WaitForTransaction();
                    }

                    if (!string.IsNullOrEmpty(control.Value))
                    {
                        SetLookUpByValue(driver, control, index);
                    }
                    else if (control.Value == "")
                    {
                        SetLookupByIndex(driver, control, index);
                    }
                    else if (control.Value == null)
                    {
                        throw new InvalidOperationException($"No value has been provided for the LookupItem {control.Name}. Please provide a value or an empty string and try again.");
                    }
                }

                input?.SendKeys(Keys.Escape); // IE wants to keep the flyout open on multi-value fields, this makes sure it closes

                return true;
            });
        }

        private void SetLookUpByValue(IWebDriver driver, LookupItem control, int index)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(d => d.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldNoRecordsText].Replace("[NAME]", control.Name) + "|" +
                AppElements.Xpath[AppReference.Entity.LookupFieldResultList].Replace("[NAME]", control.Name))));

            driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupMenu].Replace("[NAME]", control.Name)));
            var flyoutDialog = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupMenu].Replace("[NAME]", control.Name)));

            driver.WaitForTransaction();

            var lookupResultsItems = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldResultListItem].Replace("[NAME]", control.Name)));

            if (lookupResultsItems == null)
                throw new NotFoundException($"No Results Matching {control.Value} Were Found.");

            var dialogItems = OpenDialog(flyoutDialog).Value;

            driver.WaitForTransaction();
            if (dialogItems.Count == 0)
                throw new InvalidOperationException($"List does not contain a record with the name:  {control.Value}");

            if (index + 1 > dialogItems.Count)
                throw new InvalidOperationException($"List does not contain {index + 1} records. Please provide an index value less than {dialogItems.Count} ");

            var dialogItem = dialogItems[index];
            driver.ClickWhenAvailable(By.Id(dialogItem.Id));

            driver.WaitForTransaction();
        }

        private void SetLookupByIndex(IWebDriver driver, LookupItem control, int index)
        {
            driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Entity.LookupResultsDropdown].Replace("[NAME]", control.Name)));
            var lookupResultsDialog = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.LookupResultsDropdown].Replace("[NAME]", control.Name)));

            driver.WaitForTransaction();
            driver.WaitFor(d => d.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldResultListItem].Replace("[NAME]", control.Name))).Count > 0);
            var lookupResults = LookupResultsDropdown(lookupResultsDialog).Value;

            driver.WaitForTransaction();
            if (lookupResults.Count == 0)
                throw new InvalidOperationException($"No results exist in the Recently Viewed flyout menu. Please provide a text value for {control.Name}");

            if (index + 1 > lookupResults.Count)
                throw new InvalidOperationException($"Recently Viewed list does not contain {index + 1} records. Please provide an index value less than {lookupResults.Count}");

            var lookupResult = lookupResults[index];
            driver.ClickWhenAvailable(By.Id(lookupResult.Id));

            driver.WaitForTransaction();
        }

        /// <summary>
        /// Sets the value of a picklist or status field.
        /// </summary>
        /// <param name="option">The option you want to set.</param>
        /// <example>xrmApp.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });</example>
        public BrowserCommandResult<bool> SetValue(OptionSet option)
        {
            return Execute(GetOptions($"Set OptionSet Value: {option.Name}"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));

                if (fieldContainer.FindElements(By.TagName("select")).Count > 0)
                {
                    var select = fieldContainer.FindElement(By.TagName("select"));
                    var options = select.FindElements(By.TagName("option"));

                    foreach (var op in options)
                    {
                        if (op.Text != option.Value && op.GetAttribute("value") != option.Value) continue;
                        op.Click();
                        break;
                    }
                }
                else if (fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusCombo].Replace("[NAME]", option.Name))).Count > 0)
                {
                    // This is for statuscode (type = status) that should act like an optionset doesn't doesn't follow the same pattern when rendered
                    driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusComboButton].Replace("[NAME]", option.Name)));

                    var listBox = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusComboList].Replace("[NAME]", option.Name)));
                    var options = listBox.FindElements(By.TagName("li"));

                    foreach (var op in options)
                    {
                        if (op.Text != option.Value && op.GetAttribute("value") != option.Value) continue;
                        op.Click();
                        break;
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Field: {option.Name} Does not exist");
                }
                return true;
            });
        }

        /// <summary>
        /// Sets the value of a Boolean Item.
        /// </summary>
        /// <param name="option">The boolean field name.</param>
        /// <example>xrmApp.Entity.SetValue(new BooleanItem { Name = "donotemail", Value = true });</example>
        public BrowserCommandResult<bool> SetValue(BooleanItem option)
        {
            return Execute(GetOptions($"Set BooleanItem Value: {option.Name}"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));

                var hasRadio = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioContainer].Replace("[NAME]", option.Name)));
                var hasCheckbox = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));
                var hasList = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));

                if (hasRadio)
                {
                    var trueRadio = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioTrue].Replace("[NAME]", option.Name)));
                    var falseRadio = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioFalse].Replace("[NAME]", option.Name)));

                    if (option.Value && bool.Parse(falseRadio.GetAttribute("aria-checked")) || !option.Value && bool.Parse(trueRadio.GetAttribute("aria-checked")))
                    {
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioContainer].Replace("[NAME]", option.Name)));
                    }
                }
                else if (hasCheckbox)
                {
                    var checkbox = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));

                    if (option.Value && !checkbox.Selected || !option.Value && checkbox.Selected)
                    {
                        driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckboxContainer].Replace("[NAME]", option.Name)));
                    }
                }
                else if (hasList)
                {
                    var list = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));
                    var options = list.FindElements(By.TagName("option"));
                    var selectedOption = options.FirstOrDefault(a => a.HasAttribute("data-selected") && bool.Parse(a.GetAttribute("data-selected")));
                    var unselectedOption = options.FirstOrDefault(a => !a.HasAttribute("data-selected"));

                    var trueOptionSelected = false;
                    if (selectedOption != null)
                    {
                        trueOptionSelected = selectedOption.GetAttribute("value") == "1";
                    }

                    if (option.Value && !trueOptionSelected || !option.Value && trueOptionSelected)
                    {
                        if (unselectedOption != null)
                        {
                            driver.ClickWhenAvailable(By.Id(unselectedOption.GetAttribute("id")));
                        }
                    }
                }
                else
                    throw new InvalidOperationException($"Field: {option.Name} Does not exist");


                return true;
            });
        }

        /// <summary>
        /// Sets the value of a Date Field.
        /// </summary>
        /// <param name="field">Date field name.</param>
        /// <param name="date">DateTime value.</param>
        /// <param name="format">Datetime format matching Short Date & Time formatting personal options.</param>
        /// <example>xrmApp.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));</example>
        public BrowserCommandResult<bool> SetValue(string field, DateTime date, string format = "M/d/yyyy h:mm tt")
        {
            return Execute(GetOptions($"Set Value: {field}"), driver =>
            {
                driver.WaitForTransaction();

                var dateField = AppElements.Xpath[AppReference.Entity.FieldControlDateTimeInputUCI].Replace("[FIELD]", field);

                if (driver.HasElement(By.XPath(dateField)))
                {
                    var fieldElement = driver.ClickWhenAvailable(By.XPath(dateField));
                    fieldElement.Click();
                    if (fieldElement.GetAttribute("value").Length > 0)
                    {
                        fieldElement.SendKeys(Keys.Control + "a");
                        fieldElement.SendKeys(Keys.Backspace);

                        var timefields = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.FieldControlDateTimeTimeInputUCI].Replace("[FIELD]", field)));
                        if (timefields.Any())
                        {
                            driver.ClearFocus();
                            driver.WaitForTransaction();
                        }
                    }
                    ThinkTime(1000);
                    var value = date.ToString(format);
                    fieldElement.SendKeys(value);

                    try
                    {
                        driver.WaitFor(d => fieldElement.GetAttribute("value") == value);
                    }
                    catch (WebDriverTimeoutException ex)
                    {
                        throw new InvalidOperationException($"Timeout after 30 seconds. Expected: {value}. Actual: {fieldElement.GetAttribute("value")}", ex);
                    }
                    driver.ClearFocus();
                }
                else
                    throw new InvalidOperationException($"Field: {field} Does not exist");

                return true;
            });
        }

        /// <summary>
        /// Sets/Removes the value from the multselect type control
        /// </summary>
        /// <param name="option">Object of type MultiValueOptionSet containing name of the Field and the values to be set/removed</param>
        /// <param name="removeExistingValues">False - Values will be set. True - Values will be removed</param>
        /// <returns>True on success</returns>
        internal BrowserCommandResult<bool> SetValue(MultiValueOptionSet option, bool removeExistingValues = false)
        {
            return Execute(GetOptions($"Set MultiValueOptionSet Value: {option.Name}"), driver =>
            {
                if (removeExistingValues)
                {
                    RemoveMultiOptions(option);
                }
                else
                {
                    AddMultiOptions(option);
                }
                return true;
            });
        }

        /// <summary>
        /// Removes the value from the multselect type control
        /// </summary>
        /// <param name="option">Object of type MultiValueOptionSet containing name of the Field and the values to be removed</param>
        /// <returns></returns>
        private BrowserCommandResult<bool> RemoveMultiOptions(MultiValueOptionSet option)
        {
            return Execute(GetOptions($"Remove Multi Select Value: {option.Name}"), driver =>
            {
                string xpath = AppElements.Xpath[AppReference.MultiSelect.SelectedRecord].Replace("[NAME]", Elements.ElementId[option.Name]);
                // If there is already some pre-selected items in the div then we must determine if it
                // actually exists and simulate a set focus event on that div so that the input textbox
                // becomes visible.
                var listItems = driver.FindElements(By.XPath(xpath));
                if (listItems.Any())
                {
                    listItems.First().SendKeys("");
                }

                // If there are large number of options selected then a small expand collapse 
                // button needs to be clicked to expose all the list elements.
                xpath = AppElements.Xpath[AppReference.MultiSelect.ExpandCollapseButton].Replace("[NAME]", Elements.ElementId[option.Name]);
                var expandCollapseButtons = driver.FindElements(By.XPath(xpath));
                if (expandCollapseButtons.Any())
                {
                    expandCollapseButtons.First().Click(true);
                }

                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.MultiSelect.InputSearch].Replace("[NAME]", Elements.ElementId[option.Name])));
                foreach (var optionValue in option.Values)
                {
                    xpath = string.Format(AppElements.Xpath[AppReference.MultiSelect.SelectedRecordButton].Replace("[NAME]", Elements.ElementId[option.Name]), optionValue);
                    var listItemObjects = driver.FindElements(By.XPath(xpath));
                    var loopCounts = listItemObjects.Any() ? listItemObjects.Count : 0;

                    for (int i = 0; i < loopCounts; i++)
                    {
                        // With every click of the button, the underlying DOM changes and the
                        // entire collection becomes stale, hence we only click the first occurance of
                        // the button and loop back to again find the elements and anyother occurance
                        driver.FindElements(By.XPath(xpath)).First().Click(true);
                    }
                }
                return true;
            });
        }

        /// <summary>
        /// Sets the value from the multselect type control
        /// </summary>
        /// <param name="option">Object of type MultiValueOptionSet containing name of the Field and the values to be set</param>
        /// <returns></returns>
        private BrowserCommandResult<bool> AddMultiOptions(MultiValueOptionSet option)
        {
            return Execute(GetOptions($"Add Multi Select Value: {option.Name}"), driver =>
            {
                string xpath = AppElements.Xpath[AppReference.MultiSelect.SelectedRecord].Replace("[NAME]", Elements.ElementId[option.Name]);
                // If there is already some pre-selected items in the div then we must determine if it
                // actually exists and simulate a set focus event on that div so that the input textbox
                // becomes visible.
                var listItems = driver.FindElements(By.XPath(xpath));
                if (listItems.Any())
                {
                    listItems.First().SendKeys("");
                }

                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.MultiSelect.InputSearch].Replace("[NAME]", Elements.ElementId[option.Name])));
                foreach (var optionValue in option.Values)
                {
                    xpath = string.Format(AppElements.Xpath[AppReference.MultiSelect.FlyoutList].Replace("[NAME]", Elements.ElementId[option.Name]), optionValue);
                    var flyout = driver.FindElements(By.XPath(xpath));
                    if (flyout.Any())
                    {
                        flyout.First().Click(true);
                    }
                }

                // Click on the div containing textbox so that the floyout collapses or else the flyout
                // will interfere in finding the next multiselect control which by chance will be lying
                // behind the flyout control.
                //driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.MultiSelect.DivContainer].Replace("[NAME]", Elements.ElementId[option.Name])));
                xpath = AppElements.Xpath[AppReference.MultiSelect.DivContainer].Replace("[NAME]", Elements.ElementId[option.Name]);
                var divElements = driver.FindElements(By.XPath(xpath));
                if (divElements.Any())
                {
                    divElements.First().Click(true);
                }
                return true;
            });
        }

        internal BrowserCommandResult<Field> GetField(string field)
        {
            return Execute(GetOptions($"Get Field"), driver =>
            {
                Field returnField = new Field(driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", field))));
                returnField.Name = field;

                return returnField;

            });
        }

        internal BrowserCommandResult<string> GetValue(string field)
        {
            return Execute(GetOptions($"Get Value"), driver =>
            {
                string text = string.Empty;
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", field)));

                if (fieldContainer.FindElements(By.TagName("input")).Count > 0)
                {
                    var input = fieldContainer.FindElement(By.TagName("input"));
                    if (input != null)
                    {
                        IWebElement fieldValue = input.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldValue].Replace("[NAME]", field)));
                        text = fieldValue.GetAttribute("value").ToString();

                        // Needed if getting a date field which also displays time as there isn't a date specifc GetValue method
                        var timefields = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.FieldControlDateTimeTimeInputUCI].Replace("[FIELD]", field)));
                        if (timefields.Any())
                        {
                            text = $" {timefields.First().GetAttribute("value")}";
                        }
                    }
                }
                else if (fieldContainer.FindElements(By.TagName("textarea")).Count > 0)
                {
                    text = fieldContainer.FindElement(By.TagName("textarea")).GetAttribute("value");
                }
                else
                {
                    throw new Exception($"Field with name {field} does not exist.");
                }

                return text;
            });
        }

        /// <summary>
        /// Gets the value of a Lookup.
        /// </summary>
        /// <param name="control">The lookup field name of the lookup.</param>
        /// <example>xrmApp.Entity.GetValue(new Lookup { Name = "primarycontactid" });</example>
        public BrowserCommandResult<string> GetValue(LookupItem control)
        {
            return Execute($"Get Lookup Value: {control.Name}", driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupFieldContainer].Replace("[NAME]", control.Name)));

                bool found;
                string lookupValue;
                try
                {
                    found = TryGetLookupValue(fieldContainer, out lookupValue);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Field: {control.Name} Does not exist", ex);
                }
                if (!found)
                    throw new InvalidOperationException($"Field: {control.Name} Does not exist");

                return lookupValue;
            });
        }

        private static bool TryGetLookupValue(IWebElement fieldContainer, out string lookupValue)
        {
            var input = fieldContainer.FindElements(By.TagName("input")).FirstOrDefault();
            if (input != null)
            {
                lookupValue = input.GetAttribute("value");
                return true;
            }

            var label = fieldContainer.FindElements(By.XPath(".//label")).FirstOrDefault();
            if (label != null)
            {
                lookupValue = label.GetAttribute("innerText");
                return true;
            }
            lookupValue = null;
            return false;
        }

        /// <summary>
        /// Gets the value of an ActivityParty Lookup.
        /// </summary>
        /// <param name="control">The lookup field name of the lookup.</param>
        /// <example>xrmApp.Entity.GetValue(new LookupItem[] { new LookupItem { Name = "to" } });</example>
        public BrowserCommandResult<string[]> GetValue(LookupItem[] controls)
        {
            return Execute($"Get ActivityParty Lookup Value: {controls.First().Name}", driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupFieldContainer].Replace("[NAME]", controls.First().Name)));

                var existingValues = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExistingValue].Replace("[NAME]", controls.First().Name)));

                var expandCollapseButtons = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExpandCollapseButton].Replace("[NAME]", controls.First().Name)));
                if (expandCollapseButtons.Any())
                {
                    expandCollapseButtons.First().Click(true);

                    driver.WaitFor(x => x.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExistingValue].Replace("[NAME]", controls.First().Name))).Count > existingValues.Count);
                    existingValues = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExistingValue].Replace("[NAME]", controls.First().Name)));
                }

                string[] lookupValues = null;

                try
                {
                    if (existingValues.Any())
                    {
                        char[] trimCharacters = { '', '\r', '\n', '', '', '' }; //IE can return line breaks
                        lookupValues = existingValues.Select(v => v.GetAttribute("innerText").Trim(trimCharacters)).ToArray();
                    }
                    else if (fieldContainer.FindElements(By.TagName("input")).Any())
                    {
                        lookupValues = null;
                    }
                    else
                    {
                        throw new InvalidOperationException($"Field: {controls.First().Name} Does not exist");
                    }
                }
                catch (Exception exp)
                {
                    throw new InvalidOperationException($"Field: {controls.First().Name} Does not exist", exp);
                }

                return lookupValues;
            });
        }

        /// <summary>
        /// Gets the value of a picklist or status field.
        /// </summary>
        /// <param name="option">The option you want to set.</param>
        /// <example>xrmApp.Entity.GetValue(new OptionSet { Name = "preferredcontactmethodcode"}); </example>
        internal BrowserCommandResult<string> GetValue(OptionSet option)
        {
            return Execute($"Get OptionSet Value: {option.Name}", driver =>
            {
                var text = string.Empty;
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));

                if (fieldContainer.FindElements(By.TagName("select")).Count > 0)
                {
                    var select = fieldContainer.FindElement(By.TagName("select"));
                    var options = select.FindElements(By.TagName("option"));
                    foreach (var op in options)
                    {
                        if (!op.Selected) continue;
                        text = op.Text;
                        break;
                    }
                }
                else if (fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusCombo].Replace("[NAME]", option.Name))).Count > 0)
                {
                    // This is for statuscode (type = status) that should act like an optionset doesn't doesn't follow the same pattern when rendered
                    var valueSpan = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityOptionsetStatusTextValue].Replace("[NAME]", option.Name)));

                    text = valueSpan.Text;
                }
                else
                {
                    throw new InvalidOperationException($"Field: {option.Name} Does not exist");
                }
                return text;

            });
        }

        /// <summary>
        /// Sets the value of a Boolean Item.
        /// </summary>
        /// <param name="option">The boolean field name.</param>
        /// <example>xrmApp.Entity.GetValue(new BooleanItem { Name = "creditonhold" });</example>
        internal BrowserCommandResult<bool> GetValue(BooleanItem option)
        {
            return Execute($"Get BooleanItem Value: {option.Name}", driver =>
            {
                var check = false;

                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));

                var hasRadio = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioContainer].Replace("[NAME]", option.Name)));
                var hasCheckbox = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));
                var hasList = fieldContainer.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));

                if (hasRadio)
                {
                    var trueRadio = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldRadioTrue].Replace("[NAME]", option.Name)));

                    check = bool.Parse(trueRadio.GetAttribute("aria-checked"));
                }
                else if (hasCheckbox)
                {
                    var checkbox = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldCheckbox].Replace("[NAME]", option.Name)));

                    check = bool.Parse(checkbox.GetAttribute("aria-checked"));
                }
                else if (hasList)
                {
                    var list = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityBooleanFieldList].Replace("[NAME]", option.Name)));
                    var options = list.FindElements(By.TagName("option"));
                    var selectedOption = options.FirstOrDefault(a => a.HasAttribute("data-selected") && bool.Parse(a.GetAttribute("data-selected")));

                    if (selectedOption != null)
                    {
                        check = int.Parse(selectedOption.GetAttribute("value")) == 1;
                    }
                }
                else
                    throw new InvalidOperationException($"Field: {option.Name} Does not exist");

                return check;
            });
        }

        /// <summary>
        /// Gets the value from the multselect type control
        /// </summary>
        /// <param name="option">Object of type MultiValueOptionSet containing name of the Field</param>
        /// <returns>MultiValueOptionSet object where the values field contains all the contact names</returns>
        internal BrowserCommandResult<MultiValueOptionSet> GetValue(MultiValueOptionSet option)
        {
            return Execute(GetOptions($"Get Multi Select Value: {option.Name}"), driver =>
            {
                // If there are large number of options selected then a small expand collapse 
                // button needs to be clicked to expose all the list elements.
                string xpath = AppElements.Xpath[AppReference.MultiSelect.ExpandCollapseButton].Replace("[NAME]", Elements.ElementId[option.Name]);
                var expandCollapseButtons = driver.FindElements(By.XPath(xpath));
                if (expandCollapseButtons.Any())
                {
                    expandCollapseButtons.First().Click(true);
                }

                var returnValue = new MultiValueOptionSet();
                returnValue.Name = option.Name;

                xpath = AppElements.Xpath[AppReference.MultiSelect.SelectedRecordLabel].Replace("[NAME]", Elements.ElementId[option.Name]);
                var labelItems = driver.FindElements(By.XPath(xpath));
                if (labelItems.Any())
                {
                    returnValue.Values = labelItems.Select(x => x.Text).ToArray();
                }

                return returnValue;
            });
        }

        /// <summary>
        /// Returns the ObjectId of the entity
        /// </summary>
        /// <returns>Guid of the Entity</returns>
        internal BrowserCommandResult<Guid> GetObjectId(int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Get Object Id"), driver =>
            {
                var objectId = driver.ExecuteScript("return Xrm.Page.data.entity.getId();");

                Guid oId;
                if (!Guid.TryParse(objectId.ToString(), out oId))
                    throw new NotFoundException("Unable to retrieve object Id for this entity");

                return oId;
            });
        }

        /// <summary>
        /// Returns the Entity Name of the entity
        /// </summary>
        /// <returns>Entity Name of the Entity</returns>
        internal BrowserCommandResult<string> GetEntityName(int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Get Entity Name"), driver =>
            {
                var entityName = driver.ExecuteScript("return Xrm.Page.data.entity.getEntityName();").ToString();

                if (string.IsNullOrEmpty(entityName))
                {
                    throw new NotFoundException("Unable to retrieve Entity Name for this entity");
                }

                return entityName;
            });
        }

        internal BrowserCommandResult<List<GridItem>> GetSubGridItems(string subgridName)
        {
            return Execute(GetOptions($"Get Subgrid Items for Subgrid {subgridName}"), driver =>
            {
                List<GridItem> subGridRows = new List<GridItem>();

                if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridTitle].Replace("[NAME]", subgridName))))
                    throw new NotFoundException($"{subgridName} subgrid not found. Subgrid names are case sensitive.  Please make sure casing is the same.");

                //Find the subgrid contents
                var subGrid = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridContents].Replace("[NAME]", subgridName)));

                //Find the columns
                List<string> columns = new List<string>();

                var headerCells = subGrid.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridHeaders]));

                foreach (IWebElement headerCell in headerCells)
                {
                    columns.Add(headerCell.Text);
                }

                //Find the rows
                var rows = subGrid.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridRows]));

                //Process each row
                foreach (IWebElement row in rows)
                {
                    List<string> cellValues = new List<string>();
                    GridItem item = new GridItem();

                    //Get the entityId and entity Type
                    if (row.GetAttribute("data-lp-id") != null)
                    {
                        var rowAttributes = row.GetAttribute("data-lp-id").Split('|');
                        item.Id = Guid.Parse(rowAttributes[3]);
                        item.EntityName = rowAttributes[4];
                    }

                    var cells = row.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridCells]));

                    if (cells.Count > 0)
                    {
                        foreach (IWebElement thisCell in cells)
                            cellValues.Add(thisCell.Text);

                        for (int i = 0; i < columns.Count; i++)
                        {
                            //The first cell is always a checkbox for the record.  Ignore the checkbox.
                            item[columns[i]] = cellValues[i + 1];
                        }

                        subGridRows.Add(item);
                    }
                }

                return subGridRows;
            });
        }

        internal BrowserCommandResult<bool> OpenSubGridRecord(string subgridName, int index = 0)
        {
            return Execute(GetOptions($"Open Subgrid record for subgrid { subgridName}"), driver =>
            {
                //Find the Grid
                var subGrid = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridContents].Replace("[NAME]", subgridName)));

                //Get the GridName
                string subGridName = subGrid.GetAttribute("data-id").Replace("dataSetRoot_", string.Empty);

                //cell-0 is the checkbox for each record
                var checkBox = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.SubGridRecordCheckbox].Replace("[INDEX]", index.ToString()).Replace("[NAME]", subGridName)));

                driver.DoubleClick(checkBox);

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<int> GetSubGridItemsCount(string subgridName)
        {
            return Execute(GetOptions($"Get Subgrid Items Count for subgrid { subgridName}"), driver =>
            {
                List<GridItem> rows = GetSubGridItems(subgridName);
                return rows.Count;

            });
        }

        /// <summary>
        /// Click the magnifying glass icon for the lookup control supplied
        /// </summary>
        /// <param name="control">The LookupItem field on the form</param>
        /// <returns></returns>
        internal BrowserCommandResult<bool> SelectLookup(LookupItem control)
        {
            return Execute(GetOptions($"Select Lookup Field {control.Name}"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.FieldLookupButton].Replace("[NAME]", control.Name))))
                {
                    var lookupButton = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.FieldLookupButton].Replace("[NAME]", control.Name)));

                    lookupButton.Hover(driver);

                    driver.WaitForTransaction();

                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.SearchButtonIcon])).Click(true);
                }
                else
                    throw new NotFoundException($"Lookup field {control.Name} not found");

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<string> GetHeaderValue(LookupItem control)
        {
            return Execute(GetOptions($"Get Header LookupItem Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(control);
            });
        }

        internal BrowserCommandResult<string[]> GetHeaderValue(LookupItem[] controls)
        {
            return Execute(GetOptions($"Get Header Activityparty LookupItem Value {controls.First().Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(controls);
            });
        }

        internal BrowserCommandResult<string> GetHeaderValue(string control)
        {
            return Execute(GetOptions($"Get Header Value {control}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(control);
            });
        }

        internal BrowserCommandResult<MultiValueOptionSet> GetHeaderValue(MultiValueOptionSet control)
        {
            return Execute(GetOptions($"Get Header MultiValueOptionSet Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(control);
            });
        }

        internal BrowserCommandResult<string> GetHeaderValue(OptionSet control)
        {
            return Execute(GetOptions($"Get Header OptionSet Value {control}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(control);
            });
        }

        internal BrowserCommandResult<bool> GetHeaderValue(BooleanItem control)
        {
            return Execute(GetOptions($"Get Header BooleanItem Value {control}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                return GetValue(control);
            });
        }

        internal BrowserCommandResult<string> GetStatusFromFooter()
        {
            return Execute(GetOptions($"Get Status value from footer"), driver =>
            {
                if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityFooter])))
                    throw new NotFoundException("Unable to find footer on the form");

                var footer = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityFooter]));

                var status = footer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.FooterStatusValue]));

                if (string.IsNullOrEmpty(status.Text))
                    return "unknown";

                return status.Text;

            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(string field, string value)
        {
            return Execute(GetOptions($"Set Header Value {field}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(field, value);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(LookupItem control)
        {
            return Execute(GetOptions($"Set Header LookupItem Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(control);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(LookupItem[] controls)
        {
            return Execute(GetOptions($"Set Header Activityparty LookupItem Value {controls[0].Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(controls);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(MultiValueOptionSet control)
        {
            return Execute(GetOptions($"Set Header MultiValueOptionSet Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(control);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(OptionSet control)
        {
            return Execute(GetOptions($"Set Header OptionSet Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(control);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(BooleanItem control)
        {
            return Execute(GetOptions($"Set Header BooleanItem Value {control.Name}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(control);
                ThinkTime(1000);
                return true;
            });
        }

        internal BrowserCommandResult<bool> SetHeaderValue(string field, DateTime date, string format)
        {
            return Execute(GetOptions($"Set Header Value {field}"), driver =>
            {
                TryExpandHeaderFlyout(driver);
                SetValue(field, date, format);
                ThinkTime(1000);
                return true;
            });
        }

        internal void TryExpandHeaderFlyout(IWebDriver driver)
        {
            if (!driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.EntityHeader])))
                throw new NotFoundException("Unable to find header on the form");

            var xPath = By.XPath(AppElements.Xpath[AppReference.Entity.HeaderFlyoutButton]);
            bool expanded = bool.Parse(driver.FindElement(xPath).GetAttribute("aria-expanded"));

            if (!expanded)
                driver.ClickWhenAvailable(xPath);

            ThinkTime(1000);
        }

        internal BrowserCommandResult<bool> ClearValue(string fieldName)
        {
            return Execute(GetOptions($"Clear Field {fieldName}"), driver =>
            {
                SetValue(fieldName, string.Empty);

                return true;
            });
        }

        internal BrowserCommandResult<bool> ClearValue(LookupItem control, bool removeAll = true)
        {
            return Execute(GetOptions($"Clear Field {control.Name}"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupFieldContainer].Replace("[NAME]", control.Name)));

                fieldContainer.Hover(driver);

                var existingValues = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldDeleteExistingValue].Replace("[NAME]", control.Name)));

                var expandCollapseButtons = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExpandCollapseButton].Replace("[NAME]", control.Name)));
                if (expandCollapseButtons.Any())
                {
                    expandCollapseButtons.First().Click(true);

                    driver.WaitFor(x => x.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldExistingValue].Replace("[NAME]", control.Name))).Count > existingValues.Count);
                }
                else
                {
                    if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldHoverExistingValue].Replace("[NAME]", control.Name))))
                    {
                        var existingList = fieldContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldHoverExistingValue].Replace("[NAME]", control.Name)));
                        existingList.SendKeys(Keys.Clear);
                    }
                }
                driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldLookupSearchButton].Replace("[NAME]", control.Name)));

                existingValues = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldDeleteExistingValue].Replace("[NAME]", control.Name)));
                if (existingValues.Count == 0)
                    return true;

                if (removeAll)
                {
                    // Removes all selected items
                    while (existingValues.Count > 0)
                    {
                        driver.ClickWhenAvailable(By.Id(existingValues.First().GetAttribute("id")));
                        existingValues = fieldContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.Entity.LookupFieldDeleteExistingValue].Replace("[NAME]", control.Name)));
                    }
                }
                else
                {
                    // Removes an individual item by value or index
                    if (!string.IsNullOrEmpty(control.Value))
                    {
                        foreach (var existingValue in existingValues)
                        {
                            if (existingValue.GetAttribute("aria-label").EndsWith(control.Value))
                            {
                                driver.ClickWhenAvailable(By.Id(existingValue.GetAttribute("id")));
                                return true;
                            }
                        }

                        throw new InvalidOperationException($"Field '{control.Name}' does not contain a record with the name:  {control.Value}");
                    }
                    else if (control.Value == "")
                    {
                        if (control.Index + 1 > existingValues.Count)
                            throw new InvalidOperationException($"Field '{control.Name}' does not contain {control.Index + 1} records. Please provide an index value less than {existingValues.Count}");

                        driver.ClickWhenAvailable(By.Id(existingValues[control.Index].GetAttribute("id")));
                    }
                    else if (control.Value == null)
                    {
                        throw new InvalidOperationException($"No value or index has been provided for the LookupItem {control.Name}. Please provide an value or an empty string or an index and try again.");
                    }
                }

                return true;
            });
        }

        internal BrowserCommandResult<bool> ClearValue(OptionSet control)
        {
            return Execute(GetOptions($"Clear Field {control.Name}"), driver =>
            {
                control.Value = "-1";
                SetValue(control);

                return true;
            });
        }

        internal BrowserCommandResult<bool> ClearValue(MultiValueOptionSet control)
        {
            return Execute(GetOptions($"Clear Field {control.Name}"), driver =>
            {
                RemoveMultiOptions(control);

                return true;
            });
        }

        internal BrowserCommandResult<bool> SelectForm(string formName)
        {
            return Execute(GetOptions($"Select Form {formName}"), driver =>
            {
                driver.WaitForTransaction();

                if (!driver.HasElement(By.XPath(Elements.Xpath[Reference.Entity.FormSelector])))
                    throw new NotFoundException("Unable to find form selector on the form");

                var formSelector = driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Entity.FormSelector]));
                // Click didn't work with IE
                formSelector.SendKeys(Keys.Enter);

                driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Entity.FormSelectorFlyout]));

                var flyout = driver.FindElement(By.XPath(Elements.Xpath[Reference.Entity.FormSelectorFlyout]));
                var forms = flyout.FindElements(By.XPath(Elements.Xpath[Reference.Entity.FormSelectorItem]));

                var form = forms.FirstOrDefault(a => a.GetAttribute("data-text").EndsWith(formName, StringComparison.OrdinalIgnoreCase));
                if (form == null)
                    throw new NotFoundException($"Form {formName} is not in the form selector");

                driver.ClickWhenAvailable(By.Id(form.GetAttribute("id")));

                driver.WaitForPageToLoad();
                driver.WaitUntilClickable(By.XPath(Elements.Xpath[Reference.Entity.Form]),
                    new TimeSpan(0, 0, 30),
                    null,
                    d => { throw new Exception("CRM Record is Unavailable or not finished loading. Timeout Exceeded"); }
                );

                return true;
            });
        }

        internal BrowserCommandResult<bool> AddValues(LookupItem[] controls, int index = 0)
        {
            return Execute(GetOptions($"Add values {controls.First().Name}"), driver =>
            {
                SetValue(controls, index, false);

                return true;
            });
        }

        internal BrowserCommandResult<bool> RemoveValues(LookupItem[] controls)
        {
            return Execute(GetOptions($"Remove values {controls.First().Name}"), driver =>
            {
                foreach (var control in controls)
                {
                    ClearValue(control);
                }

                return true;
            });
        }

        #endregion

        #region Lookup 
        internal BrowserCommandResult<bool> OpenLookupRecord(int index)
        {
            return Execute(GetOptions("Select Lookup Record"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Lookup.LookupResultRows])))
                {
                    var rows = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Lookup.LookupResultRows]));

                    if (rows.Count > 0)
                        rows.FirstOrDefault().Click(true);
                }
                else
                    throw new NotFoundException("No rows found");

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> SearchLookupField(LookupItem control, string searchCriteria)
        {
            return Execute(GetOptions("Search Lookup Record"), driver =>
            {
                //Click in the field and enter values
                control.Value = searchCriteria;
                SetValue(control);

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> SelectLookupRelatedEntity(string entityName)
        {
            //Click the Related Entity on the Lookup Flyout
            return Execute(GetOptions($"Select Lookup Related Entity {entityName}"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Lookup.RelatedEntityLabel].Replace("[NAME]", entityName))))
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Lookup.RelatedEntityLabel].Replace("[NAME]", entityName))).Click(true);
                else
                    throw new NotFoundException($"Lookup Entity {entityName} not found");

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> SwitchLookupView(string viewName)
        {
            return Execute(GetOptions($"Select Lookup View {viewName}"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Lookup.ChangeViewButton])))
                {
                    //Click Change View 
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Lookup.ChangeViewButton])).Click(true);

                    driver.WaitForTransaction();

                    //Click View Requested 
                    var rows = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.Lookup.ViewRows]));
                    if (rows.Any(x => x.Text.Equals(viewName, StringComparison.OrdinalIgnoreCase)))
                        rows.First(x => x.Text.Equals(viewName, StringComparison.OrdinalIgnoreCase)).Click(true);
                    else
                        throw new NotFoundException($"View {viewName} not found");
                }

                else
                    throw new NotFoundException("Lookup menu not visible");

                driver.WaitForTransaction();
                return true;
            });

            return true;
        }

        internal BrowserCommandResult<bool> SelectLookupNewButton()
        {
            return Execute(GetOptions("Click New Lookup Button"), driver =>
            {
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.Lookup.NewButton])))
                {
                    var newButton = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Lookup.NewButton]));

                    if (newButton.GetAttribute("disabled") == null)
                        driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Lookup.NewButton])).Click();
                    else
                        throw new ElementNotInteractableException("New button is not enabled.  If this is a mulit-entity lookup, please use SelectRelatedEntity first.");
                }
                else
                    throw new NotFoundException("New button not found.");

                driver.WaitForTransaction();

                return true;
            });
        }

        #endregion

        #region Timeline

        /// <summary>
        /// This method opens the popout menus in the Dynamics 365 pages. 
        /// This method uses a thinktime since after the page loads, it takes some time for the 
        /// widgets to load before the method can find and popout the menu.
        /// </summary>
        /// <param name="popoutName">The By Object of the Popout menu</param>
        /// <param name="popoutItemName">The By Object of the Popout Item name in the popout menu</param>
        /// <param name="thinkTime">Amount of time(milliseconds) to wait before this method will click on the "+" popout menu.</param>
        /// <returns>True on success, False on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> OpenAndClickPopoutMenu(By menuName, By menuItemName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute($"Open menu", driver =>
            {
                driver.ClickWhenAvailable(menuName);
                try
                {
                    driver.WaitUntilAvailable(menuItemName);
                    driver.ClickWhenAvailable(menuItemName);
                }
                catch
                {
                    // Element is stale reference is thrown here since the HTML components 
                    // get destroyed and thus leaving the references null. 
                    // It is expected that the components will be destroyed and the next 
                    // action should take place after it and hence it is ignored.
                    return false;
                }
                return true;
            });
        }
        internal BrowserCommandResult<bool> Delete(int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Delete Entity"), driver =>
            {
                var deleteBtn = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.Delete]),
    "Delete Button is not available");

                deleteBtn?.Click();
                ConfirmationDialog(true);

                driver.WaitForTransaction();

                return true;
            });
        }
        internal BrowserCommandResult<bool> Assign(string userOrTeamToAssign, int thinkTime = Constants.DefaultThinkTime)
        {
            //Click the Assign Button on the Entity Record
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Assign Entity"), driver =>
            {
                var assignBtn = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.Assign]),
                    "Assign Button is not available");

                assignBtn?.Click();
                AssignDialog(Dialogs.AssignTo.User, userOrTeamToAssign);

                return true;
            });
        }
        internal BrowserCommandResult<bool> SwitchProcess(string processToSwitchTo, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Switch BusinessProcessFlow"), driver =>
            {
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Entity.ProcessButton]), new TimeSpan(0, 0, 5));
                var processBtn = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.ProcessButton]));
                processBtn?.Click();

                try
                {
                    driver.WaitUntilAvailable(
                        By.XPath(AppElements.Xpath[AppReference.Entity.SwitchProcess]),
                        new TimeSpan(0, 0, 5),
                        d => { driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.SwitchProcess])); },
                        d => { throw new InvalidOperationException("The Switch Process Button is not available."); }
                            );
                }
                catch (StaleElementReferenceException)
                {
                    Console.WriteLine("ignoring stale element exceptions");
                }
                //switchProcessBtn?.Click();

                SwitchProcessDialog(processToSwitchTo);

                return true;
            });
        }
        internal BrowserCommandResult<bool> CloseOpportunity(bool closeAsWon, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);
            string xPathQuery = string.Empty;

            if (closeAsWon)
                xPathQuery = AppElements.Xpath[AppReference.Entity.CloseOpportunityWin];
            else
                xPathQuery = AppElements.Xpath[AppReference.Entity.CloseOpportunityLoss];

            return Execute(GetOptions($"Close Opportunity"), driver =>
            {
                var closeBtn = driver.WaitUntilAvailable(By.XPath(xPathQuery), "Opportunity Close Button is not available");

                closeBtn?.Click();
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.Dialogs.CloseOpportunity.Ok]));
                CloseOpportunityDialog(true);

                return true;
            });
        }
        internal BrowserCommandResult<bool> CloseOpportunity(double revenue, DateTime closeDate, string description, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Close Opportunity"), driver =>
            {
                SetValue(Elements.ElementId[AppReference.Dialogs.CloseOpportunity.ActualRevenueId], revenue.ToString());
                SetValue(Elements.ElementId[AppReference.Dialogs.CloseOpportunity.CloseDateId], closeDate);
                SetValue(Elements.ElementId[AppReference.Dialogs.CloseOpportunity.DescriptionId], description);

                driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Dialogs.CloseOpportunity.Ok]),
                    new TimeSpan(0, 0, 5),
                    d => { driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Dialogs.CloseOpportunity.Ok])); },
                    d => { throw new InvalidOperationException("The Close Opportunity dialog is not available."); });

                return true;
            });
        }

        /// <summary>
        /// This method opens the popout menus in the Dynamics 365 pages. 
        /// This method uses a thinktime since after the page loads, it takes some time for the 
        /// widgets to load before the method can find and popout the menu.
        /// </summary>
        /// <param name="popoutName">The name of the Popout menu</param>
        /// <param name="popoutItemName">The name of the Popout Item name in the popout menu</param>
        /// <param name="thinkTime">Amount of time(milliseconds) to wait before this method will click on the "+" popout menu.</param>
        /// <returns>True on success, False on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> OpenAndClickPopoutMenu(string popoutName, string popoutItemName, int thinkTime = Constants.DefaultThinkTime)
        {
            return OpenAndClickPopoutMenu(By.XPath(Elements.Xpath[popoutName]), By.XPath(Elements.Xpath[popoutItemName]), thinkTime);
        }


        /// <summary>
        /// Provided a By object which represents a HTML Button object, this method will
        /// find it and click it.
        /// </summary>
        /// <param name="by">The object of Type By which represents a HTML Button object</param>
        /// <returns>True on success, False/Exception on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> ClickButton(By by)
        {
            return Execute($"Open Timeline Add Post Popout", driver =>
            {
                var button = driver.WaitUntilAvailable(by);
                if (button.TagName.Equals("button"))
                {
                    try
                    {
                        driver.ClickWhenAvailable(by);
                    }
                    catch
                    {
                        // Element is stale reference is thrown here since the HTML components 
                        // get destroyed and thus leaving the references null. 
                        // It is expected that the components will be destroyed and the next 
                        // action should take place after it and hence it is ignored.
                    }
                    return true;
                }
                else if (button.FindElements(By.TagName("button")).Any())
                {
                    button.FindElements(By.TagName("button")).First().Click();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException($"Control does not exist");
                }
            });
        }

        /// <summary>
        /// Provided a fieldname as a XPath which represents a HTML Button object, this method will
        /// find it and click it.
        /// </summary>
        /// <param name="fieldNameXpath">The field as a XPath which represents a HTML Button object</param>
        /// <returns>True on success, Exception on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> ClickButton(string fieldNameXpath)
        {
            try
            {
                return ClickButton(By.XPath(fieldNameXpath));
            }
            catch (Exception e)
            {
                throw new InvalidOperationException($"Field: {fieldNameXpath} with Does not exist", e);
            }
        }

        /// <summary>
        /// Generic method to help click on any item which is clickable or uniquely discoverable with a By object.
        /// </summary>
        /// <param name="by">The xpath of the HTML item as a By object</param>
        /// <returns>True on success, Exception on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> SelectTab(string tabName, string subTabName = "", int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute($"Select Tab", driver =>
            {
                IWebElement tabList = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TabList]));

                ClickTab(tabList, AppElements.Xpath[AppReference.Entity.Tab], tabName);

                //Click Sub Tab if provided
                if (!string.IsNullOrEmpty(subTabName))
                {
                    ClickTab(tabList, AppElements.Xpath[AppReference.Entity.SubTab], subTabName);
                }

                driver.WaitForTransaction();
                return true;
            });
        }

        internal void ClickTab(IWebElement tabList, string xpath, string name)
        {
            // Look for the tab in the tab list, else in the more tabs menu
            IWebElement searchScope = null;
            var xpathByName = By.XPath(string.Format(xpath, name));
            if (tabList.HasElement(xpathByName))
                searchScope = tabList;
            else if (tabList.TryFindElement(By.XPath(AppElements.Xpath[AppReference.Entity.MoreTabs]), out IWebElement moreTabsButton))
            {
                moreTabsButton.Click();
                searchScope = Browser.Driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.MoreTabsMenu]));
            }

            if (searchScope != null && searchScope.TryFindElement(xpathByName, out IWebElement listItem))
                listItem.Click(true);
            else
                throw new Exception($"The tab with name: {name} does not exist");
        }

        /// <returns>True on success, Exception on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> IsTabVisible(string tabName, string subTabName = "", int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute("Is Tab Visible", driver =>
            {
                IWebElement tabList = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TabList]));
                if (tabList == null)
                    return false;

                bool result;
                //Click Sub Tab if provided
                if (string.IsNullOrEmpty(subTabName))
                {
                    result = IsTabVisible(tabList, AppElements.Xpath[AppReference.Entity.Tab], tabName);
                }
                else
                {
                    ClickTab(tabList, AppElements.Xpath[AppReference.Entity.Tab], tabName);
                    result = IsTabVisible(tabList, AppElements.Xpath[AppReference.Entity.SubTab], subTabName);
                }
                return result;
            });
        }


        internal bool IsTabVisible(IWebElement tabList, string xpath, string name)
        {
            // Look for the tab in the tab list, else in the more tabs menu
            IWebElement searchScope = null;
            if (tabList.HasElement(By.XPath(string.Format(xpath, name))))
            {
                searchScope = tabList;

            }
            else if (tabList.TryFindElement(By.XPath(AppElements.Xpath[AppReference.Entity.MoreTabs]), out IWebElement moreTabsButton))
            {
                moreTabsButton.Click();
                searchScope = Browser.Driver.FindElement(By.XPath(AppElements.Xpath[AppReference.Entity.MoreTabsMenu]));
            }

            IWebElement listItem = null;
            var result = searchScope != null && searchScope.TryFindElement(By.XPath(string.Format(xpath, name)), out listItem);
            return result && listItem != null;
        }

        /// <summary>
        /// A generic setter method which will find the HTML Textbox/Textarea object and populate
        /// it with the value provided. The expected tag name is to make sure that it hits
        /// the expected tag and not some other object with the similar fieldname.
        /// </summary>
        /// <param name="fieldName">The name of the field representing the HTML Textbox/Textarea object</param>
        /// <param name="value">The string value which will be populated in the HTML Textbox/Textarea</param>
        /// <param name="expectedTagName">Expected values - textbox/textarea</param>
        /// <returns>True on success, Exception on failure to invoke any action</returns>
        internal BrowserCommandResult<bool> SetValue(string fieldName, string value, string expectedTagName)
        {
            return Execute($"SetValue (Generic)", driver =>
            {
                var inputbox = driver.WaitUntilAvailable(By.XPath(Elements.Xpath[fieldName]));
                if (expectedTagName.Equals(inputbox.TagName, StringComparison.InvariantCultureIgnoreCase))
                {
                    inputbox.Click();
                    inputbox.Clear();
                    inputbox.SendKeys(value);
                    return true;
                }
                else
                {
                    throw new InvalidOperationException($"Field: {fieldName} with tagname {expectedTagName} Does not exist");
                }
            });
        }

        #endregion

        #region BusinessProcessFlow

        /// <summary>
        /// Set Value
        /// </summary>
        /// <param name="field">The field</param>
        /// <param name="value">The value</param>
        /// <example>xrmApp.BusinessProcessFlow.SetValue("firstname", "Test");</example>
        internal BrowserCommandResult<bool> BPFSetValue(string field, string value)
        {
            return Execute(GetOptions($"Set BPF Value"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.TextFieldContainer].Replace("[NAME]", field)));

                if (fieldContainer.FindElements(By.TagName("input")).Count > 0)
                {
                    var input = fieldContainer.FindElement(By.TagName("input"));
                    if (input != null)
                    {
                        input.Click(true);
                        input.Clear();
                        input.SendKeys(value, true);
                        input.SendKeys(Keys.Tab);
                    }
                }
                else if (fieldContainer.FindElements(By.TagName("textarea")).Count > 0)
                {
                    fieldContainer.FindElement(By.TagName("textarea")).Click();
                    fieldContainer.FindElement(By.TagName("textarea")).Clear();
                    fieldContainer.FindElement(By.TagName("textarea")).SendKeys(value);
                }
                else
                {
                    throw new Exception($"Field with name {field} does not exist.");
                }

                return true;
            });
        }

        /// <summary>
        /// Sets the value of a picklist.
        /// </summary>
        /// <param name="option">The option you want to set.</param>
        /// <example>xrmBrowser.BusinessProcessFlow.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });</example>
        public BrowserCommandResult<bool> BPFSetValue(OptionSet option)
        {
            return Execute(GetOptions($"Set BPF Value: {option.Name}"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.Entity.TextFieldContainer].Replace("[NAME]", option.Name)));

                if (fieldContainer.FindElements(By.TagName("select")).Count > 0)
                {
                    var select = fieldContainer.FindElement(By.TagName("select"));
                    var options = select.FindElements(By.TagName("option"));

                    foreach (var op in options)
                    {
                        if (op.Text != option.Value && op.GetAttribute("value") != option.Value) continue;
                        op.Click(true);
                        break;
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Field: {option.Name} Does not exist");
                }
                return true;
            });
        }

        /// <summary>
        /// Sets the value of a Boolean Item.
        /// </summary>
        /// <param name="option">The option you want to set.</param>
        /// <example>xrmBrowser.BusinessProcessFlow.SetValue(new BooleanItem { Name = "preferredcontactmethodcode"});</example>
        public BrowserCommandResult<bool> BPFSetValue(BooleanItem option)
        {
            return Execute(GetOptions($"Set BPF Value: {option.Name}"), driver =>
            {
                var fieldContainer = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.BooleanFieldContainer].Replace("[NAME]", option.Name)));
                if (!option.Value)
                {
                    if (!fieldContainer.Selected)
                    {
                        fieldContainer.Click(true);
                    }
                    else
                    {
                        fieldContainer.Click(true);
                    }
                }
                else
                {
                    if (fieldContainer.Selected)
                    {
                        fieldContainer.Click(true);
                    }
                    else
                    {
                        fieldContainer.Click(true);
                    }
                }
                return true;
            });
        }

        /// <summary>
        /// Sets the value of a Date Field.
        /// </summary>
        /// <param name="field">The field id or name.</param>
        /// <param name="date">DateTime value.</param>
        /// <param name="format">DateTime format</param>
        /// <example> xrmBrowser.BusinessProcessFlow.SetValue("birthdate", DateTime.Parse("11/1/1980"));</example>
        public BrowserCommandResult<bool> BPFSetValue(string field, DateTime date, string format = "MM dd yyyy")
        {
            return Execute(GetOptions($"Set BPF Value: {field}"), driver =>
            {
                var dateField = AppElements.Xpath[AppReference.BusinessProcessFlow.DateTimeFieldContainer].Replace("[FIELD]", field);

                if (driver.HasElement(By.XPath(dateField)))
                {
                    var fieldElement = driver.ClickWhenAvailable(By.XPath(dateField));

                    if (fieldElement.GetAttribute("value").Length > 0)
                    {
                        //fieldElement.Click();
                        //fieldElement.SendKeys(date.ToString(format));
                        //fieldElement.SendKeys(Keys.Enter);

                        fieldElement.Click();
                        Browser.ThinkTime(250);
                        fieldElement.Click();
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(date.ToString(format), true);
                        Browser.ThinkTime(500);
                        fieldElement.SendKeys(Keys.Tab);
                        Browser.ThinkTime(250);
                    }
                    else
                    {
                        fieldElement.Click();
                        Browser.ThinkTime(250);
                        fieldElement.Click();
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Backspace);
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(date.ToString(format));
                        Browser.ThinkTime(250);
                        fieldElement.SendKeys(Keys.Tab);
                        Browser.ThinkTime(250);
                    }
                }
                else
                    throw new InvalidOperationException($"Field: {field} Does not exist");

                return true;
            });
        }

        internal BrowserCommandResult<bool> NextStage(string stageName, Field businessProcessFlowField = null, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Next Stage"), driver =>
            {
                //Find the Business Process Stages
                var processStages = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.NextStage_UCI]));

                if (processStages.Count == 0)
                    return true;

                foreach (var processStage in processStages)
                {
                    var labels = processStage.FindElements(By.TagName("label"));

                    //Click the Label of the Process Stage if found
                    foreach (var label in labels)
                    {
                        if (label.Text.Equals(stageName, StringComparison.OrdinalIgnoreCase))
                        {
                            label.Click();
                        }
                    }
                }

                var flyoutFooterControls = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.Flyout_UCI]));

                foreach (var control in flyoutFooterControls)
                {
                    //If there's a field to enter, fill it out
                    if (businessProcessFlowField != null)
                    {
                        var bpfField = control.FindElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.BusinessProcessFlowFieldName].Replace("[NAME]", businessProcessFlowField.Name)));

                        if (bpfField != null)
                        {
                            bpfField.Click();
                            for (int i = 0; i < businessProcessFlowField.Value.Length; i++)
                            {
                                bpfField.SendKeys(businessProcessFlowField.Value.Substring(i, 1));
                            }
                        }
                    }

                    //Click the Next Stage Button
                    var nextButton = control.FindElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.NextStageButton]));
                    nextButton.Click();
                }

                return true;
            });
        }

        internal BrowserCommandResult<bool> SelectStage(string stageName, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Select Stage: {stageName}"), driver =>
            {

                //Find the Business Process Stages
                var processStages = driver.FindElements(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.NextStage_UCI]));

                foreach (var processStage in processStages)
                {
                    var labels = processStage.FindElements(By.TagName("label"));

                    //Click the Label of the Process Stage if found
                    foreach (var label in labels)
                    {
                        if (label.Text.Equals(stageName, StringComparison.OrdinalIgnoreCase))
                        {
                            label.Click();
                        }
                    }
                }

                driver.WaitForTransaction();

                return true;
            });
        }

        internal BrowserCommandResult<bool> SetActive(string stageName = "", int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Set Active Stage: {stageName}"), driver =>
            {
                if (!string.IsNullOrEmpty(stageName))
                {
                    SelectStage(stageName);

                    if (!driver.HasElement(By.XPath("//button[contains(@data-id,'setActiveButton')]")))
                        throw new NotFoundException($"Unable to find the Set Active button. Please verify the stage name { stageName } is correct.");

                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.SetActiveButton])).Click(true);

                    driver.WaitForTransaction();
                }

                return true;
            });
        }

        internal BrowserCommandResult<bool> BPFPin(string stageName, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Pin BPF: {stageName}"), driver =>
            {
                //Click the BPF Stage
                SelectStage(stageName, 0);
                driver.WaitForTransaction();

                //Pin the Stage
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.PinStageButton])))
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.PinStageButton])).Click();
                else
                    throw new NotFoundException($"Pin button for stage {stageName} not found.");

                driver.WaitForTransaction();
                return true;
            });
        }

        internal BrowserCommandResult<bool> BPFClose(string stageName, int thinkTime = Constants.DefaultThinkTime)
        {
            return Execute(GetOptions($"Close BPF: {stageName}"), driver =>
            {
                //Click the BPF Stage
                SelectStage(stageName, 0);
                driver.WaitForTransaction();

                //Pin the Stage
                if (driver.HasElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.CloseStageButton])))
                    driver.FindElement(By.XPath(AppElements.Xpath[AppReference.BusinessProcessFlow.CloseStageButton])).Click(true);
                else
                    throw new NotFoundException($"Close button for stage {stageName} not found.");

                driver.WaitForTransaction();
                return true;
            });
        }

        #endregion

        #region GlobalSearch
        /// <summary>
        /// Searches for the specified criteria in Global Search.
        /// </summary>
        /// <param name="criteria">Search criteria.</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param> time.</param>
        /// <example>xrmBrowser.GlobalSearch.Search("Contoso");</example>
        internal BrowserCommandResult<bool> GlobalSearch(string criteria, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Global Search: {criteria}"), driver =>
            {
                var xpathToButton = By.XPath(AppElements.Xpath[AppReference.Navigation.SearchButton]);
                IWebElement openSearchButton =  driver.WaitUntilClickable(xpathToButton, new TimeSpan(0, 0, 5),
                                            failureCallback: d => throw new InvalidOperationException("The Global Search button is not available."));
                openSearchButton.Click();

                var xpathToTextBox = By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Text]);
                var textbox = driver.WaitUntilVisible(xpathToTextBox, failureCallback: d => throw new InvalidOperationException("The Global Search is not available."));
                if (textbox != null)
                {
                    var searchType = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Type])).GetAttribute("value");
                    IWebElement button = null;
                    if (searchType == "1") //Categorized Search
                    {
                        button = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.CategorizedSearchButton]));
                    }
                    else if (searchType == "0") //Relevance Search
                    {
                        button = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.RelevanceSearchButton]));
                    }

                    if (button != null)
                    {
                        textbox.SendKeys(criteria, true);
                        button.Click(true);
                    }
                    else
                    {
                        throw new InvalidOperationException("The Global Search text field is not available.");
                    }
                }
                return true;
            });
        }

        /// <summary>
        /// Filter by entity in the Global Search Results.
        /// </summary>
        /// <param name="entity">The entity you want to filter with.</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param>
        /// <example>xrmBrowser.GlobalSearch.FilterWith("Account");</example>
        public BrowserCommandResult<bool> FilterWith(string entity, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Filter With: {entity}"), driver =>
            {
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Filter]),
                                        new TimeSpan(0, 0, 10),
                                        e =>
                                        {
                                            var options = e.FindElements(By.TagName("option"));
                                            e.Click();

                                            IWebElement option = options.FirstOrDefault(x => x.Text == entity);

                                            if (option == null)
                                                throw new InvalidOperationException($"Entity '{entity}' does not exist in the Filter options.");

                                            option.Click();
                                        },
                                        f => throw new InvalidOperationException("Filter With picklist is not available. The timeout period elapsed waiting for the picklist to be available."));

                return true;
            });
        }

        /// <summary>
        /// Filter by group and value in the Global Search Results.
        /// </summary>
        /// <param name="filterby">The Group that contains the filter you want to use.</param>
        /// <param name="value">The value listed in the group by area.</param>
        /// <example>xrmBrowser.GlobalSearch.Filter("Record Type", "Accounts");</example>
        public BrowserCommandResult<bool> Filter(string filterBy, string value, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Filter With: {value}"), driver =>
            {
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.GroupContainer].Replace("[NAME]", filterBy)),
                                        new TimeSpan(0, 0, 10),
                                        e =>
                                        {
                                            var groupContainer = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.GroupContainer].Replace("[NAME]", filterBy)));
                                            var filter = groupContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.FilterValue].Replace("[NAME]", value)));

                                            if (filter == null)
                                                throw new InvalidOperationException($"Filter By Value '{value}' does not exist in the Filter options.");

                                            filter.Click();
                                        },
                                        f => throw new InvalidOperationException("Filter With picklist is not available. The timeout period elapsed waiting for the picklist to be available."));

                return true;
            });
        }

        /// <summary>
        /// Opens the specified record in the Global Search Results.
        /// </summary>
        /// <param name="entity">The entity you want to open a record.</param>
        /// <param name="index">The index of the record you want to open.</param>
        /// <param name="thinkTime">Used to simulate a wait time between human interactions. The Default is 2 seconds.</param> time.</param>
        /// <example>xrmBrowser.GlobalSearch.OpenRecord("Accounts",0);</example>
        public BrowserCommandResult<bool> OpenGlobalSearchRecord(string entity, int index, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Open Global Search Record"), driver =>
            {
                var searchType = driver.WaitUntilAvailable(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Type])).GetAttribute("value");

                if (searchType == "1") //Categorized Search
                {
                    driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Container]),
                                            Constants.DefaultTimeout,
                                            null,
                                            d => throw new InvalidOperationException("Search Results is not available"));


                    var resultsContainer = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Container]));

                    var entityContainer = resultsContainer.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.EntityContainer].Replace("[NAME]", entity)));

                    if (entityContainer == null)
                        throw new InvalidOperationException($"Entity {entity} was not found in the results");

                    var records = entityContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Records]));

                    if (records == null)
                        throw new InvalidOperationException($"No records found for entity {entity}");

                    records[index].Click();
                    driver.WaitUntilClickable(By.XPath(AppElements.Xpath[AppReference.Entity.Form]),
                        new TimeSpan(0, 0, 30),
                        null,
                        d => { throw new Exception("CRM Record is Unavailable or not finished loading. Timeout Exceeded"); }
                    );
                }
                else if (searchType == "0")   //Relevance Search
                {
                    var resultsContainer = driver.FindElement(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.RelevanceResultsContainer]));
                    var records = resultsContainer.FindElements(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.RelevanceResults].Replace("[ENTITY]", entity.ToUpper())));

                    if (records.Count >= index + 1)
                        records[index].Click(true);
                }

                return true;
            });
        }


        /// <summary>
        /// Changes the search type used for global search
        /// </summary>
        /// <param name="type">The type of search that you want to do.</param>
        /// <example>xrmBrowser.GlobalSearch.ChangeSearchType("Categorized");</example>
        public BrowserCommandResult<bool> ChangeSearchType(string type, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Change Search Type"), driver =>
            {
                driver.WaitUntilVisible(By.XPath(AppElements.Xpath[AppReference.GlobalSearch.Type]),
                                        Constants.DefaultTimeout,
                                        e =>
                                        {
                                            var options = e.FindElements(By.TagName("option"));
                                            var option = options.FirstOrDefault(x => x.Text.Trim() == type);
                                            if (option == null)
                                                return;

                                            e.Click(true);
                                            option.Click(true);
                                        },
                                        d => throw new InvalidOperationException("Search Results is not available"));
                return true;
            });
        }
        #endregion

        #region Dashboard
        internal BrowserCommandResult<bool> SelectDashboard(string dashboardName, int thinkTime = Constants.DefaultThinkTime)
        {
            Browser.ThinkTime(thinkTime);

            return Execute(GetOptions($"Select Dashboard"), driver =>
            {
                //Click the drop-down arrow
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Dashboard.DashboardSelector]));
                //Select the dashboard
                driver.ClickWhenAvailable(By.XPath(AppElements.Xpath[AppReference.Dashboard.DashboardItemUCI].Replace("[NAME]", dashboardName)));

                return true;
            });
        }
        #endregion

        #region PerformanceCenter
        internal void EnablePerformanceCenter()
        {
            Browser.Driver.Navigate().GoToUrl($"{Browser.Driver.Url}&perf=true");
            Browser.Driver.WaitForPageToLoad();
            Browser.Driver.WaitForTransaction();
        }
        #endregion

        internal void ThinkTime(int milliseconds)
        {
            Browser.ThinkTime(milliseconds);
        }
        internal void Dispose()
        {
            Browser.Dispose();
        }
    }
}
