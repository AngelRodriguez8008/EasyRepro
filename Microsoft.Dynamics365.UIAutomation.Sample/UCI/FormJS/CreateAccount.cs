using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api.UCI.Extensions;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using Guid = System.Guid;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI.FormJS
{
    [TestClass]
    public class CreateAccount
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        // Allow trigger to complete the set value action
        private string _enter(string value) => value + Keys.Enter;
        private string _timed(string value) => $"{value} {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}";
        private Uri _create(string entityType) => new Uri(_xrmUri, $"/main.aspx?etn={entityType}&pagetype=entityrecord");

        
        [TestMethod]
        public void CreateNewAccount()
        {
            var options = TestSettings.Options;
            options.TimeFactor = 0.5f;
            var client = new WebClient(options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
             
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                
                xrmApp.Grid.SwitchView("All Accounts");
                
                xrmApp.CommandBar.ClickCommand("New");
                
                xrmApp.Entity.SetValue("name", _timed("Test API Account"));
                xrmApp.Entity.SetValue("telephone1", "555-555-5555");
                //xrmApp.Entity.SetValue(new BooleanItem { Name =  "creditonhold", Value = true });

                xrmApp.Entity.Save();
                // Is not possible to check Save Success or Fail
            }
        }

        [TestMethod]
        public void CreateNewAccount_ReduceTheTime()
        {
            var options = TestSettings.Options;
            options.TimeFactor = 0.5f;
            var client = new WebClient(options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
             
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.ThinkTime(500);
                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.ThinkTime(2000);
                xrmApp.Grid.SwitchView("Active Accounts");

                xrmApp.ThinkTime(1000);
                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.ThinkTime(5000);
                xrmApp.Entity.SetValue("name", "Test API Account");
                xrmApp.Entity.SetValue("telephone1", "555-555-5555");
                xrmApp.Entity.SetValue(new BooleanItem { Name =  "creditonhold", Value = true });

                xrmApp.CommandBar.ClickCommand("Save & Close");
                // Is not possible to check Save Success or Fail
                xrmApp.ThinkTime(2000);
            }
        }

        [TestMethod]
        public void CreateNewAccount_InconclusiveSetValue()
        {
            var options = TestSettings.Options;
            options.TimeFactor = 0.5f;
            var client = new WebClient(options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);
            
                xrmApp.ThinkTime(5000);
                xrmApp.Entity.SetValue("name", "Test API Account");
                var expectedPhone1 = "555-555-5555";
                xrmApp.Entity.SetValue("telephone1", expectedPhone1);

                xrmApp.ThinkTime(2000);
            
                string phone = xrmApp.FormJS.GetAttributeValue<string>("telephone1");
                Assert.AreEqual(expectedPhone1, phone);
             }   
        }

        [TestMethod]
        public void CreateNewAccount_ExecuteJavaScript_Generic()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.ThinkTime(2000);
                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.ThinkTime(5000);
                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                xrmApp.ThinkTime(1000);

                string code = "var a = 1+1; return a";
                var result = client.Browser.Driver.ExecuteJavaScript<long>(code);
                
                Assert.AreEqual(2, result);
                xrmApp.ThinkTime(5000);
            }
        }
        
        [TestMethod]
        public void CreateNewAccount_ExecuteJavaScript_Xrm()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                xrmApp.ThinkTime(2000);
                var name = _timed($"Test Account");
                xrmApp.Entity.SetValue("name", _enter(name));

                xrmApp.ThinkTime(1000);
                
                object result = xrmApp.Entity.GetValue("name");
                Assert.AreEqual(name, result);

                string code = @"
                    var result = Xrm.Page.getAttribute('name').getValue(); 
                    return result
                    ";
                
                //xrmApp.Entity.SwitchToContentFrame();
                result = client.Browser.Driver.ExecuteJavaScript<string>(code);

                Assert.AreEqual(name, result);
                xrmApp.ThinkTime(5000);
            }
        }

        [TestMethod]
        public void CreateNewAccount_ExecuteJavaScript_Xrm_GetId()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                 xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);
               

                xrmApp.ThinkTime(2000);
                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", _enter(name));
                
                string code = @"return Xrm.Page.data.entity.getId()";
                
                //xrmApp.Entity.SwitchToContentFrame();
                string result = client.Browser.Driver.ExecuteJavaScript<string>(code);

                Assert.IsNotNull(result);
                Assert.AreEqual(string.Empty, result);
                xrmApp.ThinkTime(5000);
            }
        }
        [TestMethod]
        public void CreateNewAccount_ExecuteJavaScript_Xrm_Using_HelperClass()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                xrmApp.ThinkTime(1000);
                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", _enter(name));

                xrmApp.ThinkTime(500);

                string result = xrmApp.Entity.GetValue("name");
                Assert.AreEqual(name, result);

                result = xrmApp.FormJS.GetAttributeValue<string>("name");
                Assert.AreEqual(name, result);
                xrmApp.ThinkTime(5000);
            }
        }

        [TestMethod]
        public void CreateNewAccount_CheckSave()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                xrmApp.ThinkTime(5000);
                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                Guid id = xrmApp.FormJS.GetEntityId();
                Assert.AreEqual(default(Guid), id);

                xrmApp.Entity.Save();
                xrmApp.ThinkTime(5000);

                id = xrmApp.FormJS.GetEntityId();
                Assert.AreNotEqual(default(Guid), id);
                Console.WriteLine(id);
            }
        }

        [TestMethod]
        public void CreateNewAccount_SetPhoneNummer()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                string expectedPhone1 = "+43 122466";
                xrmApp.Entity.SetValue("telephone1", _enter(expectedPhone1));
                xrmApp.ThinkTime(500);

                string expectedPhone2 = "+43 1223344";
                xrmApp.Entity.SetValue("fax", _enter(expectedPhone2));
                xrmApp.ThinkTime(500);

                string phone = xrmApp.Entity.GetValue("telephone1");
                Assert.AreEqual(expectedPhone1, phone);
                
                phone = xrmApp.Entity.GetValue("fax");
                Assert.AreEqual(expectedPhone2, phone);

                phone = xrmApp.FormJS.GetAttributeValue<string>("telephone1");
                Assert.AreEqual(expectedPhone1, phone);
                
                phone = xrmApp.FormJS.GetAttributeValue<string>("fax");
                Assert.AreEqual(expectedPhone2, phone);
            }
        }
        

        // Using SetValue to Override a Text field don't work
        [TestMethod]
        public void CreateNewAccount_OverridePhoneNummer_Fail()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                string expectedPhone1 = "+43 122466";
                xrmApp.Entity.SetValue("telephone1", _enter(expectedPhone1));
                xrmApp.ThinkTime(500);

                string expectedPhone2 = "+43 1223344";
                xrmApp.Entity.SetValue("fax", _enter(expectedPhone2));
                xrmApp.ThinkTime(500);

                string phone = xrmApp.Entity.GetValue("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.Entity.GetValue("fax");
                Assert.AreEqual(expectedPhone2, phone);

                expectedPhone1 = "+43 122466 22";
                xrmApp.Entity.SetValue("telephone1", _enter(expectedPhone1));
                xrmApp.ThinkTime(500);

                expectedPhone2 = "+43 1223344 22";
                xrmApp.Entity.SetValue("fax", _enter(expectedPhone2));
                xrmApp.ThinkTime(500);

                phone = xrmApp.Entity.GetValue("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.Entity.GetValue("fax");
                Assert.AreEqual(expectedPhone2, phone);

                phone = xrmApp.FormJS.GetAttributeValue<string>("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.FormJS.GetAttributeValue<string>("fax");
                Assert.AreEqual(expectedPhone2, phone);
            }
        }

        // Using SetValue to Override a Text field don't work, clear the old value solve the problem
        // There is not clear function, clear the value is not easy to simulate without FormJS
        [TestMethod]
        public void CreateNewAccount_OverridePhoneNummer_Clearing()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                string expectedPhone1 = "+43 122466";
                xrmApp.Entity.SetValue("telephone1", _enter(expectedPhone1));
                xrmApp.ThinkTime(500);

                string expectedPhone2 = "+43 1223344";
                xrmApp.Entity.SetValue("fax", _enter(expectedPhone2));
                xrmApp.ThinkTime(500);

                string phone = xrmApp.Entity.GetValue("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.Entity.GetValue("fax");
                Assert.AreEqual(expectedPhone2, phone);

                expectedPhone1 = "+43 122466 22";
                xrmApp.FormJS.Clear("telephone1");
                xrmApp.Entity.SetValue("telephone1", _enter(expectedPhone1));
                xrmApp.ThinkTime(500);

                expectedPhone2 = "+43 1223344 22";
                xrmApp.FormJS.Clear("fax");
                xrmApp.Entity.SetValue("fax", _enter(expectedPhone2));
                xrmApp.ThinkTime(500);

                phone = xrmApp.Entity.GetValue("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.Entity.GetValue("fax");
                Assert.AreEqual(expectedPhone2, phone);

                phone = xrmApp.FormJS.GetAttributeValue<string>("telephone1");
                Assert.AreEqual(expectedPhone1, phone);

                phone = xrmApp.FormJS.GetAttributeValue<string>("fax");
                Assert.AreEqual(expectedPhone2, phone);
            }
        }

        // Please: Check that you has US Dollar & Euro Currencies in your system
        [TestMethod]
        public void CreateNewAccount_OverrideCurrency_Fail()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                var expectedCurrency = "US Dollar";
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", expectedCurrency);

                string currency = xrmApp.Entity.GetLookupValue("transactioncurrencyid");
                Assert.AreEqual(expectedCurrency, currency);
            }
        }

        [TestMethod]
        public void CreateNewAccount_OverrideCurrency_Clearing()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);

                xrmApp.FormJS.Clear("transactioncurrencyid");
                var expectedCurrency = "US Dollar";
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", expectedCurrency);
                
                string currency = xrmApp.Entity.GetLookupValue("transactioncurrencyid");
                Assert.AreEqual(expectedCurrency, currency);
            }
        }
        
        [TestMethod]
        public void OnChangeContry_UpdatePhonePrefix_OverrideCurrency_UsingDialog()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                xrmApp.ThinkTime(5000);
                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);
                
                xrmApp.FormJS.Clear("transactioncurrencyid");
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", "Euro");

                xrmApp.ThinkTime(1000);

                var expectedCurrency = "US Dollar";
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", expectedCurrency);

                xrmApp.ThinkTime(5000);

                string currency = xrmApp.Entity.GetValue("transactioncurrencyid");
                Assert.AreEqual(expectedCurrency, currency);
            }
        }

        // EntityExtension is a Combination of more Commands
        [TestMethod]
        public void CreateNewAccount_OverrideCurrency_JS_Clearing_Twice()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var newAccountUrl = _create("account");
                xrmApp.OnlineLogin.Login(newAccountUrl, _username, _password);

                var name = _timed("Test Account");
                xrmApp.Entity.SetValue("name", name);
                var expectedCurrency = "Euro";
                xrmApp.FormJS.Clear("transactioncurrencyid");
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", expectedCurrency);

                xrmApp.ThinkTime(1000);

                expectedCurrency = "US Dollar";
                xrmApp.FormJS.Clear("transactioncurrencyid");
                xrmApp.Entity.SetLookupValue("transactioncurrencyid", expectedCurrency);

                xrmApp.ThinkTime(1000);

                string currency = xrmApp.Entity.GetLookupValue("transactioncurrencyid");
                Assert.AreEqual(expectedCurrency, currency);
            }
        }
    }
}