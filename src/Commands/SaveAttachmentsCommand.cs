using System;
using System.IO;
using Microsoft.Extensions.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveOutlookAttachments
{
    internal class SaveAttachmentsCommand : IAppCommand
    {
        private readonly ILogger<ListStoresCommand> logger;
        private readonly OutlookManager outlook;

        public SaveAttachmentsCommand(ILogger<ListStoresCommand> logger, OutlookManager outlook)
        {
            this.logger = logger;
            this.outlook = outlook;
        }

        public int Run(AppContext context)
        {
            int exitCode = 0;

            try
            {
                outlook.Initialise();

                if (outlook.TrySetStore(context.Options.StoreDescriptor))
                {
                    var store = outlook.GetCurrentStore();
                    logger.LogInformation("Name=\"{DisplayName}\",Path=\"{FilePath}\",Type={ExchangeStoreType}", store.DisplayName, store.FilePath, store.ExchangeStoreType);

                    if (outlook.TrySetFolder())
                    {
                        outlook.ProcessItem = ProcessItem;

                        if (context.Options.WhatIf == false)
                        {
                            Directory.CreateDirectory(context.Options.TargetFolder);
                        }

                        outlook.ForEachAttachment(context);
                    }
                }

                context.Stats.Show();
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to save Outlook attachments");
            }

            return exitCode;
        }

        private void ProcessItem(AppContext ctx, object item)
        {
            ctx.Stats.LogStats(item);

            if (item is Outlook.MailItem mi)
            {
                ProcessMailItem(ctx, mi);
            }
        }

        private void ProcessMailItem(AppContext ctx, Outlook.MailItem item)
        {
            switch (item.MessageClass)
            {
                case "IPM.Note":
                    ProcessMailAttachments(ctx, item);
                    break;
                case "IPM.Note.SMIME.MultipartSigned":
                    ProcessMailAttachments(ctx, item);
                    break;
                default:
                    logger.LogInformation("ProcessMailItem\\Not Processed\\{MessageClass}", item.MessageClass);
                    break;
            }
        }

        private void ProcessMailAttachments(AppContext ctx, Outlook.MailItem item)
        {
            if (item.Attachments != null)
            {
                foreach (Outlook.Attachment attachment in item.Attachments)
                {
                    ProcessMailAttachment(ctx, attachment);
                }
            }
        }

        private void ProcessMailAttachment(AppContext ctx, Outlook.Attachment item)
        {
            logger.LogInformation("ProcessMailAttachment\\\"{DisplayName}\",{Class},{Type},{Size}", item.DisplayName, item.Class, item.Type, item.Size);

            switch (item.Type)
            {
                case Outlook.OlAttachmentType.olByValue:
                    _ = SaveAttachment(ctx, item);
                    break;
                case Outlook.OlAttachmentType.olEmbeddeditem:
                    ProcessEmbeddedItemAttachment(ctx, item);
                    break;
                case Outlook.OlAttachmentType.olOLE:
                    _ = SaveAttachment(ctx, item);
                    break;
                default:
                    break;
            }
        }

        private void ProcessEmbeddedItemAttachment(AppContext ctx, Outlook.Attachment item)
        {
            var path = SaveAttachment(ctx, item);

            if (string.IsNullOrEmpty(path) == false && ctx.Options.WhatIf == false)
            {
                var embeddedMessage = outlook.OpenSharedItem(path);

                ProcessItem(ctx, embeddedMessage);
            }
        }

        private string SaveAttachment(AppContext ctx, Outlook.Attachment item)
        {
            string result = null;

            if (item.Size != 0)
            {
                var fileName = item.GetFileName();
                if (string.IsNullOrEmpty(fileName) == false)
                {
                    fileName = FileHelpers.CleanFileName(fileName, '_');

                    fileName = FileHelpers.GenerateUniqueFileName(ctx.Options.TargetFolder, fileName);

                    string outputFile = Path.GetFullPath(fileName);

                    logger.LogInformation("SaveAttachment\\\"{FileName}\" => \"{OutputFileName}\"", item.FileName, outputFile);
                    if (ctx.Options.WhatIf == false)
                    {
                        item.SaveAsFile(outputFile);
                    }
                }
            }
            return result;
        }
    }
}
