using System;
using Microsoft.Dynamics365.UIAutomation.Browser;

namespace Microsoft.Dynamics365.UIAutomation.Api.Pages
{
    public class FormJS : BrowserPage
    {
        private readonly IFormJSWorker _worker;

        public FormJS(InteractiveBrowser browser)
            : base(browser)
        {
            _worker = new FormJSWorker(browser);
        }

        public FormJS(IFormJSWorker worker)
        {
            _worker = worker;
        }

        public BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter) 
            => _worker.ExecuteJS(commandName, code, converter);
        public bool ExecuteJS(string commandName, string code, params object[] args) 
            => _worker.ExecuteJS(commandName, code, args);
        public bool ExecuteJS(string code, params object[] args) 
            => ExecuteJS("Execute JS via FormJS", code, args);
        public BrowserCommandResult<T> ExecuteJS<T>(string commandName, string code) 
            => ExecuteJS<T, T>(commandName, code, result => result);
        
        public BrowserCommandResult<T> GetAttributeValue<T>(string attributte)
        {
            var commandName = $"Get Attribute Value via Form JS: {attributte}";
            string code = $"return Xrm.Page.getAttribute('{attributte}').getValue()";

            return ExecuteJS<T>(commandName, code);
        }

        public bool SetAttributeValue<T>(string attributte, T value)
        {
            var commandName = $"Set Attribute Value via Form JS: {attributte}";
            string code = $"return Xrm.Page.getAttribute('{attributte}').setValue(arguments[0])";

            return ExecuteJS(commandName, code, value);
        }

        public bool Clear(string attribute)
        {
            var commandName = $"Clear Attribute via Form JS: {attribute}";
            string code = $"return Xrm.Page.getAttribute('{attribute}').setValue(null)";

            return ExecuteJS(commandName, code);
        }

        public BrowserCommandResult<Guid> GetEntityId()
        {
            var commandName = $"Get Entity Id via Form JS";
            string code = $"return Xrm.Page.data.entity.getId()";

            return ExecuteJS<string, Guid>(commandName, code, v =>
            {
                bool success = Guid.TryParse(v, out Guid result);
                return success ? result : default(Guid);
            });
        }

        public BrowserCommandResult<bool> IsControlVisible(string attributte)
        {
            var commandName = $"Get Control Visibility via Form JS: {attributte}";
            string code = $"return Xrm.Page.getControl('{attributte}').getVisible()";

            return ExecuteJS<bool>(commandName, code);
        }

        public BrowserCommandResult<bool> IsDirty(string attributte)
        {
            var commandName = $"Get Attribute IsDirty via Form JS: {attributte}";
            string code = $"return Xrm.Page.getAttribute('{attributte}').getIsDirty()";

            return ExecuteJS<bool>(commandName, code);
        }

        public enum RequiredLevel { Unknown = 0, None, Required, Recommended }

        public BrowserCommandResult<RequiredLevel> GetRequiredLevel(string attributte)
        {
            var commandName = $"Get Attribute RequiredLevel via Form JS: {attributte}";
            string code = $"return Xrm.Page.getAttribute('{attributte}').getRequiredLevel()";
            
            return ExecuteJS<string, RequiredLevel>(commandName, code,
                v =>
                {
                    bool success =  Enum.TryParse(v, true, out RequiredLevel result);
                    return success ? result : RequiredLevel.Unknown;
                });
        }

        public bool Disable(string attributte) => Enable(attributte, false);
        public bool Enable(string attributte, bool value = true)
        {
            var commandName = $"Enable Attribute via Form JS: {attributte}";
            string code = $"return Xrm.Page.getControl('{attributte}').setDisabled({(value? "false" : "true")});";
            
            return ExecuteJS(commandName, code);
        }

        public bool LoadWebResource(string webResourceName, bool async = true)
        {
            string strAsync = async ? "true" : "false";
            var commandName = $"Load Web Resource: {webResourceName}";

            // credit: Ben John => How to load Javascript from a webresource
            // https://community.dynamics.com/365/b/leichtbewoelkt/posts/how-to-load-javascript-from-a-webresource
            string code = $@"(function() {{
                let req = new XMLHttpRequest();
                req.open('GET', Xrm.Page.context.getClientUrl() + '/api/data/v8.0/webresourceset?' +
                                ""$select=content&$filter=name eq '{webResourceName}'"", {strAsync});
                req.setRequestHeader('OData-Version', '4.0');
                req.setRequestHeader('OData-MaxVersion', '4.0');
                req.setRequestHeader('Accept', 'application/json');
                req.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
                req.onreadystatechange = function () {{
                    if (this.readyState === 4) {{
                        req.onreadystatechange = null;
                        if (this.status === 200) {{                            
                            var result = (JSON.parse(this.response)).value[0].content;
                            var script = atob(result);
                            window.eval(script);                           
                            return true;
                        }}
                        else
                            console.error(this.statusText);                        
                    }}
                    return false;
                }};
                req.send();
            }})();";
            
            return ExecuteJS(commandName, code);
        }
    }
}
