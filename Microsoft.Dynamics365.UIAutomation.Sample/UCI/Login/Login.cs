using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api.UCI.Extensions;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using Guid = System.Guid;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI.Login
{
    [TestClass]
    public class Login
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);
        private readonly SecureString _mfaSecrectKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecrectKey"].ToSecureString();

        // Allow trigger to complete the set value action
        private string _enter(string value) => value + Keys.Enter;
        private string _timed(string value) => $"{value} {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}";
        
        [TestMethod]
        public void MultiFactorLogin()
        {
            var options = TestSettings.Options;
            options.TimeFactor = 0.5f;
            var client = new WebClient(options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecrectKey);
             
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                
                xrmApp.Grid.SwitchView("All Accounts");
                
                xrmApp.CommandBar.ClickCommand("New");
                
                xrmApp.Entity.SetValue("name", _timed("Test API Account"));
                xrmApp.Entity.SetValue("telephone1", "555-555-5555");
                //xrmApp.Entity.SetValue(new BooleanItem { Name =  "creditonhold", Value = true });
            }
        }
    }
}