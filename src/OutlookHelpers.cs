using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveOutlookAttachments
{
    public static class OutlookHelpers
    {
        private static readonly Assembly _outlookAssembly = Assembly.GetAssembly(typeof(Outlook.MailItem));
        private static readonly Dictionary<string, Type> _map = new Dictionary<string, Type>();

        public static Type[] GetOutlookTypes()
        {
            return _outlookAssembly.GetTypes();
        }

        public static Type GetOutlookType(object item)
        {
            var name = TypeDescriptor.GetClassName(item);

            if (_map.TryGetValue(name, out Type result) == false)
            {
                var types = GetOutlookTypes();

                result = (from type in types
                          where type.Name.Equals(name)
                          select type).SingleOrDefault();
                if (result == default)
                {
                    result = typeof(object);
                }

                _map.Add(name, result);
            }

            return result;
        }
    }
}
