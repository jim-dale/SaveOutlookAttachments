using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveOutlookAttachments
{
    internal static class OutlookExtensions
    {
        internal static bool TryGetStore(this Outlook.NameSpace item, string descriptor, out Outlook.Store result)
        {
            bool success = false;
            result = default;

            foreach (var store in item.GetStores())
            {
                if (string.Equals(descriptor, store.DisplayName, StringComparison.CurrentCultureIgnoreCase)
                    || string.Equals(descriptor, store.FilePath, StringComparison.InvariantCultureIgnoreCase))
                {
                    result = store;
                    success = true;
                }
            }

            return success;
        }

        internal static IEnumerable<Outlook.Store> GetStores(this Outlook.NameSpace item)
        {
            foreach (var store in item.Stores.OfType<Outlook.Store>())
            {
                yield return store;
            }
        }

        internal static string GetHeaders(this Outlook.MailItem item)
        {
            var result = (string)item.PropertyAccessor.GetProperty(Constants.PR_TRANSPORT_MESSAGE_HEADERS);

            return result;
        }

        public static string GetFileName(this Outlook.Attachment item)
        {
            try
            {
                return item.FileName;
            }
            catch (COMException)
            {
            }

            return default;
        }
    }
}
