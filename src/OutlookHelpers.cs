
namespace SaveOutlookAttachments
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public static class OutlookHelpers
    {
        public const string PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x";
        public const string PT_BOOLEAN = "000B";
        public const string PT_STRING8 = "001E";
        public const string PT_UNICODE = "001F";
        public const string PR_TRANSPORT_MESSAGE_HEADERS = PROPTAG + "007D" + PT_STRING8;
        public const string PR_ATTACH_CONTENT_ID = PROPTAG + "3712" + PT_STRING8;
        public const string PR_ATTACH_HIDDEN = PROPTAG + "7FFE" + PT_BOOLEAN;

        static Assembly _outlookAssembly = Assembly.GetAssembly(typeof(Outlook.MailItem));
        static Dictionary<string, Type> _map = new Dictionary<string, Type>();

        public static Type GetOutlookType(object item)
        {
            string name = TypeDescriptor.GetClassName(item);
            if (_map.TryGetValue(name, out Type result) == false)
            {
                result = (from type in _outlookAssembly.GetTypes()
                          where type.Name.Equals(name)
                          select type).SingleOrDefault();
                if (result == null)
                {
                    result = typeof(object);
                }
                _map.Add(name, result);
            }
            return result;
        }

        public static string GetHeaders(Outlook.MailItem item)
        {
            string result = (string)item.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS);
            return result;
        }

        public static bool IsEmbeddedResource(Outlook.Attachment item)
        {
            string contentId = (string)item.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID);
            return String.IsNullOrEmpty(contentId) == false;
        }
    }
}
