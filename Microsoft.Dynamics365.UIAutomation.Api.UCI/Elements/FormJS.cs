using System;
using Microsoft.Dynamics365.UIAutomation.Browser;
using FormJSPage =  Microsoft.Dynamics365.UIAutomation.Api.Pages.FormJS;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI
{
    public class FormJS : Element
    {
        private readonly FormJSWorker _worker;  
        private readonly Pages.FormJS _page;  

        public FormJS(WebClient client) 
        {
            _worker = new FormJSWorker(client);
            _page = new Pages.FormJS(_worker);
        }
        
        public BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter) => _page.ExecuteJS(commandName, code, converter);
        public bool ExecuteJS(string code, params object[] args) => _page.ExecuteJS( code, args);
        public bool ExecuteJS(string commandName, string code, params object[] args) => _page.ExecuteJS(commandName, code, args);
        public BrowserCommandResult<T> ExecuteJS<T>(string commandName, string code) => _page.ExecuteJS<T>(commandName, code);

        public BrowserCommandResult<T> GetAttributeValue<T>(string attributte) => _page.GetAttributeValue<T>(attributte);
        public bool SetAttributeValue<T>(string attributte, T value) => _page.SetAttributeValue(attributte, value);
        public bool Clear(string attribute) => _page.Clear(attribute);
        public BrowserCommandResult<Guid> GetEntityId() => _page.GetEntityId();
        public BrowserCommandResult<bool> IsControlVisible(string attributte) => _page.IsControlVisible(attributte);
        public BrowserCommandResult<bool> IsDirty(string attributte) => _page.IsDirty(attributte);
        public BrowserCommandResult<FormJSPage.RequiredLevel> GetRequiredLevel(string attributte) => _page.GetRequiredLevel(attributte);
        public bool Enable(string attributte) => _page.Enable(attributte);
        public bool Disable(string attributte) => _page.Disable(attributte);
        public bool LoadWebResource(string webResourceName, bool async = true) => _page.LoadWebResource(webResourceName, async);
    }
}
