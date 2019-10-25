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

    }
}
