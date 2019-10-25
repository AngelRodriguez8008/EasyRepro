// Created by: Rodriguez Mustelier Angel -- 
// Modify On: 2019-10-21 04:20

using System;
using Microsoft.Dynamics365.UIAutomation.Api.Pages;
using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium.Support.Extensions;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class FormJSWorker : Element, IFormJSWorker
    {
        private readonly Pages.FormJSWorker _worker;  

        public FormJSWorker(WebClient client) 
        {
            _worker = new Pages.FormJSWorker(client.Browser);
        }

        public BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter)
        {
            if (converter == null)
                throw new ArgumentNullException(nameof(converter));

            return _worker.Execute(_worker.GetOptions(commandName), driver =>
            {
                driver.WaitForPageToLoad();

                T result = driver.ExecuteJavaScript<T>(code);
                return converter(result);
            });
        }
        
        public bool ExecuteJS(string commandName, string code, params object[] args)
        {
            return _worker.Execute(_worker.GetOptions(commandName), driver =>
            {
                driver.WaitForPageToLoad();

                driver.ExecuteJavaScript(code, args);
                return true;
            });
        }
    }
}