using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedUploadFile : PluginBase
    {
        public SharePointEmbeddedUploadFile(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedUploadFile))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var containerId = PluginParameterHelper.GetRequiredString(context, "ContainerId");
            var fileName = PluginParameterHelper.GetRequiredString(context, "FileName");
            var fileContentBase64 = PluginParameterHelper.GetRequiredString(context, "FileContentBase64");
            var folderPath = PluginParameterHelper.GetOptionalString(context, "FolderPath");
            var contentType = PluginParameterHelper.GetOptionalString(context, "ContentType");

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var result = client.UploadFile(containerId, folderPath, fileName, fileContentBase64, contentType);

                context.OutputParameters["DriveId"] = result.DriveId;
                context.OutputParameters["DriveItemId"] = result.DriveItemId;
                context.OutputParameters["Name"] = result.Name;
                context.OutputParameters["SizeInBytes"] = result.Size.ToString();
                context.OutputParameters["WebUrl"] = result.WebUrl;
            }
        }
    }
}