using System;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveOutlookAttachments
{
    class Program
    {
        static void Main(string[] args)
        {
            var cfg = ArgsProcessor.Parse(args);
            if (cfg.ShowHelp)
            {
                ArgsProcessor.ShowHelp();
            }
            else
            {
                var ctx = AppContext.GetAppContextFromConfig(cfg);

                Directory.CreateDirectory(ctx.TargetFolder);

                var application = new Outlook.Application();
                try
                {
                    ctx.Session = application.GetNamespace("MAPI");

                    ProcessStore(ctx);
                }
                catch (Exception ex)
                {
                    Trace.TraceError(ex.ToString());
                    Console.Error.WriteLine(ex.ToString());
                }
                finally
                {
                    application.Quit();
                }
                ctx.Stats.Show();
            }
        }

        static void ProcessStore(AppContext ctx)
        {
            bool removeStoreAfter = false;
            Outlook.MAPIFolder rootFolder = null;

            try
            {
                if (string.IsNullOrEmpty(ctx.SourcePst) == false)
                {
                    ctx.Session.AddStore(ctx.SourcePst);
                    removeStoreAfter = true;
                }
                rootFolder = ctx.Session.Stores[ctx.StoreName].GetRootFolder();
                if (rootFolder != null)
                {
                    ProcessFolder(ctx, rootFolder);
                }
            }
            finally
            {
                if (rootFolder != null && removeStoreAfter)
                {
                    ctx.Session.RemoveStore(rootFolder);
                }
            }
        }

        private static void ProcessFolder(AppContext ctx, Outlook.MAPIFolder folder)
        {
            ProcessFolderItems(ctx, folder);

            ProcessFolders(ctx, folder.Folders);
        }

        private static void ProcessFolders(AppContext ctx, Outlook.Folders folders)
        {
            foreach (Outlook.MAPIFolder folder in folders)
            {
                ProcessFolder(ctx, folder);
            }
        }

        private static void ProcessFolderItems(AppContext ctx, Outlook.MAPIFolder folder)
        {
            var items = folder.Items.Cast<object>();

            foreach (var item in items)
            {
                ProcessItem(ctx, item);
            }
        }

        private static void ProcessItem(AppContext ctx, object item)
        {
            ctx.Stats.LogStats(item);

            switch (item)
            {
                case Outlook.MailItem mi:
                    ProcessMailItem(ctx, mi);
                    break;
                //case Outlook.NoteItem ni:
                //    ProcessNoteItem(ctx, ni);
                //    break;
                //    case Outlook.AppointmentItem ai:
                //        break;
                //    case Outlook.ContactItem ci:
                //        break;
                //    case Outlook.MeetingItem mi:
                //        break;
                //    case Outlook.PostItem pi:
                //        break;
                default:
                    break;
            }
        }

        private static void ProcessMailItem(AppContext ctx, Outlook.MailItem item)
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
                    Trace.TraceWarning($"ProcessMailItem\\Not Processed\\{item.MessageClass}");
                    break;
            }
        }

        private static void ProcessNoteItem(AppContext ctx, Outlook.NoteItem item)
        {
            dynamic result = new ExpandoObject();

            result.MessageClass = item.MessageClass;
            result.Id = item.EntryID;
            result.Created = item.CreationTime;
            result.LastModified = item.LastModificationTime;
            result.Subject = item.Subject;
            result.Body = item.Body;

            string json = JsonConvert.SerializeObject(result, Formatting.Indented);
            Trace.WriteLine(json);
        }

        private static void ProcessMailAttachments(AppContext ctx, Outlook.MailItem item)
        {
            if (item.Attachments == null)
            {
                return;
            }

            foreach (Outlook.Attachment attachment in item.Attachments)
            {
                ProcessMailAttachment(ctx, attachment);
            }
        }

        private static void ProcessMailAttachment(AppContext ctx, Outlook.Attachment item)
        {
            switch (item.Type)
            {
                case Outlook.OlAttachmentType.olByValue:
                    ProcessByValueAttachment(ctx, item);
                    break;
                case Outlook.OlAttachmentType.olEmbeddeditem:
                    ProcessEmbeddedItemAttachment(ctx, item);
                    break;
                case Outlook.OlAttachmentType.olOLE:
                    SaveOleAttachmentToFile(ctx, item);
                    break;
                default:
                    Trace.TraceInformation($"ProcessMailAttachment\\{item.DisplayName},{item.Type}");
                    break;
            }
        }

        private static void ProcessByValueAttachment(AppContext ctx, Outlook.Attachment item)
        {
            string path = SaveAttachmentToFile(ctx, item);
        }

        private static void ProcessEmbeddedItemAttachment(AppContext ctx, Outlook.Attachment item)
        {
            string path = SaveAttachmentToFile(ctx, item);
            if (String.IsNullOrEmpty(path) == false)
            {
                var newItem = ctx.Session.OpenSharedItem(path);

                ProcessItem(ctx, newItem);
            }
        }

        private static string SaveAttachmentToFile(AppContext ctx, Outlook.Attachment item)
        {
            Trace.TraceInformation($"SaveAttachmentToFile\\{item.DisplayName},{item.Type},{item.Size}");

            string result = null;

            if (item.Size != 0)
            {
                string fileName = item.FileName;
                if (fileName != null)
                {
                    result = FileHelpers.GenerateUniqueFileName(fileName, ctx.TargetFolder);

                    Trace.TraceInformation($"SaveAttachmentToFile\\{item.FileName})=>{fileName}");
                    item.SaveAsFile(result);
                }
            }
            return result;
        }

        private static void SaveOleAttachmentToFile(AppContext ctx, Outlook.Attachment item)
        {
            Trace.TraceInformation($"SaveOleAttachmentToFile\\{item.DisplayName},{item.Type},{item.Size}");

            if (item.Size != 0)
            {
                if (OutlookHelpers.IsEmbeddedResource(item) == false)
                {
                    string fileName = Path.GetRandomFileName();
                    string suffix = $" ({item.DisplayName})";

                    string path = FileHelpers.GenerateUniqueFileName(fileName, suffix, ctx.TargetFolder);

                    Trace.TraceInformation($"SaveOleAttachmentToFile\\{fileName}");
                    item.SaveAsFile(path);
                }
                else
                {
                    Trace.TraceInformation("SaveOleAttachmentToFile\\Embedded resource ignored");
                }
            }
        }
    }
}
