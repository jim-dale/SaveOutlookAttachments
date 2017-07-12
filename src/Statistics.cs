
namespace SaveOutlookAttachments
{
    using System;
    using System.Collections.Generic;

    public class Statistics
    {
        private Dictionary<Type, TypeStatistics> _typesStats = new Dictionary<Type, TypeStatistics>();

        public void LogStats(dynamic item)
        {
            Type type = OutlookHelpers.GetOutlookType(item);

            if (_typesStats.TryGetValue(type, out TypeStatistics typeStats) == false)
            {
                typeStats = new TypeStatistics();
            }
            ++typeStats.Count;
            _typesStats[type] = typeStats;

            typeStats.LogStats(item);
        }

        public void Show()
        {
            foreach (var outlookType in _typesStats)
            {
                Console.WriteLine("Type '{0}' = {1}", outlookType.Key, outlookType.Value.Count);
                foreach (var messageClass in outlookType.Value._messageClassCounters)
                {
                    Console.WriteLine("\tMessage class '{0}' = {1}", messageClass.Key, messageClass.Value);
                }
            }
        }

        private class TypeStatistics
        {
            public long Count = 0;
            public Dictionary<string, long> _messageClassCounters = new Dictionary<string, long>();

            public void LogStats(dynamic item)
            {
                string messageClass = item.MessageClass;

                _messageClassCounters.TryGetValue(messageClass, out long count);
                _messageClassCounters[messageClass] = ++count;
            }
        }
    }
}
