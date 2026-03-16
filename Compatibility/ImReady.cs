using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace ManagedIdentityPlugin
{
    /// <summary>
    /// Custom API handler "Im ready".
    /// Input: UserId (Guid) — GUID of the system user claiming work.
    /// Output: QueueItemId (Guid) — GUID of the claimed queue item.
    /// </summary>
    public class ImReady : PluginBase
    {
        public ImReady(string unsecureConfiguration, string secureConfiguration) : base(typeof(ImReady))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
                throw new ArgumentNullException(nameof(localPluginContext));

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.PluginUserService;

            if (!context.InputParameters.Contains("UserId"))
                throw new InvalidPluginExecutionException("Input parameter 'UserId' is required.");

            if (!(context.InputParameters["UserId"] is Guid userId))
                throw new InvalidPluginExecutionException("Input parameter 'UserId' must be a Guid.");

            var queueItem = FindFirstUnworkedQueueItem(service);
            if (queueItem == null)
            {
                throw new InvalidPluginExecutionException("No available queue items without 'Worked By' were found.");
            }

            var pick = new PickFromQueueRequest
            {
                QueueItemId = queueItem.Id,
                WorkerId = userId,
                RemoveQueueItem = false
            };

            service.Execute(pick);
            context.OutputParameters["QueueItemId"] = queueItem.Id;
        }

        private static Entity FindFirstUnworkedQueueItem(IOrganizationService service)
        {
            var query = new QueryExpression("queueitem")
            {
                ColumnSet = new ColumnSet("queueitemid"),
                NoLock = true,
                PageInfo = new PagingInfo
                {
                    Count = 1,
                    PageNumber = 1,
                    ReturnTotalRecordCount = false
                }
            };

            var filter = new FilterExpression(LogicalOperator.And);
            filter.AddCondition("workerid", ConditionOperator.Null);
            filter.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria = filter;
            query.AddOrder("createdon", OrderType.Ascending);

            var results = service.RetrieveMultiple(query);
            return results.Entities.Count > 0 ? results.Entities[0] : null;
        }
    }
}