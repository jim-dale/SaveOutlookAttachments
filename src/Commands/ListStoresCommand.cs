using System;
using Microsoft.Extensions.Logging;

namespace SaveOutlookAttachments
{
    internal class ListStoresCommand : IAppCommand
    {
        private readonly ILogger<ListStoresCommand> logger;
        private readonly OutlookManager outlook;

        public ListStoresCommand(ILogger<ListStoresCommand> logger, OutlookManager outlook)
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

                foreach (var store in outlook.GetStores())
                {
                    Console.WriteLine("Name=\"{0}\",Path=\"{1}\",Type={2}", store.DisplayName, store.FilePath, store.ExchangeStoreType);
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to list Outlook stores");
            }

            return exitCode;
        }
    }
}
