using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace ManagedIdentityPlugin
{
    public class DemoContactParentTracer : PluginBase
    {
        public DemoContactParentTracer(string unsecureConfiguration, string secureConfiguration) : base(typeof(DemoContactParentTracer))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.InitiatingUserService;
            var target = (Entity)context.InputParameters["Target"];

            if (!target.Contains("parentcustomerid") || target["parentcustomerid"] == null)
            {
                localPluginContext.Trace("No parent customer specified for this contact.");
                return;
            }

            var parentRef = (EntityReference)target["parentcustomerid"];
            if (parentRef.LogicalName != "account")
            {
                localPluginContext.Trace($"Parent customer is not an account (it's a {parentRef.LogicalName}). Skipping account name lookup.");
                return;
            }

            var account = service.Retrieve("account", parentRef.Id, new ColumnSet("name"));
            var accountName = account.GetAttributeValue<string>("name");
            localPluginContext.Trace($"Parent account: {accountName} ({parentRef.Id})");
        }
    }
}