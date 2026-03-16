using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    /// <summary>
    /// Calls a custom API (PowerFx function) with the Target parameter set to the target entity's GUID.
    /// The API name is provided via unsecure configuration.
    /// </summary>
    public class CallPowerFxFunction : PluginBase
    {
        private readonly string _customApiName;

        public CallPowerFxFunction(string unsecureConfiguration, string secureConfiguration) : base(typeof(CallPowerFxFunction))
        {
            if (string.IsNullOrWhiteSpace(unsecureConfiguration))
                throw new ArgumentException("Custom API name must be provided in unsecure configuration.");

            _customApiName = unsecureConfiguration;
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.PluginUserService;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
            {
                var request = new OrganizationRequest(_customApiName);
                request["Target"] = target.Id.ToString();
                service.Execute(request);
            }
            else
            {
                throw new InvalidPluginExecutionException("Target entity reference not found in InputParameters.");
            }
        }
    }
}