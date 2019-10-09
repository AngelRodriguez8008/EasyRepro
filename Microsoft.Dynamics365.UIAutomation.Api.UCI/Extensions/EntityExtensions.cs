using Microsoft.Dynamics365.UIAutomation.Browser;

namespace Microsoft.Dynamics365.UIAutomation.Api.UCI.Extensions
{
    public static class EntityExtensions
    {
        public static BrowserCommandResult<string> GetLookupValue(this Entity entity, string attribute) =>
            entity.GetValue(new LookupItem {Name = attribute});

        public static void SetLookupValue(this Entity entity, string attribute, string value)
        {
            entity.SetValue(attribute, value);
            entity.SelectLookup(new LookupItem{ Name = attribute, Index =  0});
        }

    }
}