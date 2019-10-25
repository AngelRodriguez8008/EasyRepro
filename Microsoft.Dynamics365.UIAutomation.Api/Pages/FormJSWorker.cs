// Created by: Rodriguez Mustelier Angel -- 
// Modify On: 2019-10-21 03:59

using System;
using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;

namespace Microsoft.Dynamics365.UIAutomation.Api.Pages
{
    public class FormJSWorker : BrowserPage, IFormJSWorker
    {
        public FormJSWorker(InteractiveBrowser browser)
            : base(browser)
        { }
        
        public BrowserCommandOptions GetOptions(string commandName)
        {
            return new BrowserCommandOptions(Constants.DefaultTraceSource,
                commandName,
                0,
                0,
                null,
                true,
                typeof(NoSuchElementException), typeof(StaleElementReferenceException));
        }

        public bool SwitchToContent()
        {
            Browser.Driver.SwitchTo().DefaultContent();
            //wait for the content panel to render
            Browser.Driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Frames.ContentPanel]));

            //find the crmContentPanel and find out what the current content frame ID is - then navigate to the current content frame
            var currentContentFrame = Browser.Driver.FindElement(By.XPath(Elements.Xpath[Reference.Frames.ContentPanel]))
                .GetAttribute(Elements.ElementId[Reference.Frames.ContentFrameId]);

            Browser.Driver.SwitchTo().Frame(currentContentFrame);

            return true;
        }

        public BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter)
        {
            if (converter == null)
                throw new ArgumentNullException(nameof(converter));

            return Execute(GetOptions(commandName), driver =>
            {
                SwitchToContent();

                T result = driver.ExecuteJavaScript<T>(code);
                return converter(result);
            });
        }
        
        public bool ExecuteJS(string commandName, string code, params object[] args)
        {
            return Execute(GetOptions(commandName), driver =>
            {
                SwitchToContent();

                driver.ExecuteJavaScript(code, args);
                return true;
            });
        }
    }
}