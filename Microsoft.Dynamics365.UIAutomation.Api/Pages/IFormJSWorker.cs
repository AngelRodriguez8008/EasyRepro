// Created by: Rodriguez Mustelier Angel -- 
// Modify On: 2019-10-21 04:00

using System;
using Microsoft.Dynamics365.UIAutomation.Browser;

namespace Microsoft.Dynamics365.UIAutomation.Api.Pages
{
    public interface IFormJSWorker
    {
        bool ExecuteJS(string commandName, string code, params object[] args);
        BrowserCommandResult<TResult> ExecuteJS<T, TResult>(string commandName, string code, Func<T, TResult> converter);
    }
}