using System;
using Microsoft.Dynamics365.UIAutomation.Browser;
using FormJSPage =  Microsoft.Dynamics365.UIAutomation.Api.Pages.FormJS;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class FormJS : Element
    {
        private readonly FormJSPage _worker;  

        public FormJS(WebClient client) 
        {
            _worker = new FormJSPage(client.Browser);
        }

        public BrowserCommandOptions GetOptions(string commandName) => _worker.GetOptions(commandName);
        public bool SwitchToContent() => _worker.SwitchToContent();
        public BrowserCommandResult<T> GetAttributeValue<T>(string attributte) => _worker.GetAttributeValue<T>(attributte);
        public bool SetAttributeValue<T>(string attributte, T value) => _worker.SetAttributeValue(attributte, value);
        public bool Clear(string attribute) => _worker.Clear(attribute);
        public BrowserCommandResult<Guid> GetEntityId() => _worker.GetEntityId();
        public BrowserCommandResult<bool> IsControlVisible(string attributte) => _worker.IsControlVisible(attributte);
        public BrowserCommandResult<bool> IsDirty(string attributte) => _worker.IsDirty(attributte);
        public BrowserCommandResult<FormJSPage.RequiredLevel> GetRequiredLevel(string attributte) => _worker.GetRequiredLevel(attributte);
        public BrowserCommandResult<T> ExecuteJS<T>(string commandName, string code) => _worker.ExecuteJS<T>(commandName, code);
        public BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter) => _worker.ExecuteJS(commandName, code, converter);
        public bool ExecuteJS(string commandName, string code, params object[] args) => _worker.ExecuteJS(commandName, code, args);
    }
}
