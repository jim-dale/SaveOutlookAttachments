using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace SaveOutlookAttachments
{
    partial class Program
    {
        private static void ConfigureServices(ServiceCollection services)
        {
            services.AddLogging(configure =>
            {
                //configure.SetMinimumLevel(LogLevel.Warning);
                configure.AddConsole();
            });
            services.AddSingleton<OutlookManager>();
            services.AddTransient<ListStoresCommand>();
            services.AddTransient<SaveAttachmentsCommand>();
        }

        static int Main(string[] args)
        {
            int exitCode = 1;

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            using (var serviceProvider = serviceCollection.BuildServiceProvider())
            {
                var config = new AppConfig()
                    .UseCommandLine(args);

                if (config.ShowHelp)
                {
                    AppConfig.OutputHelp();
                    if (config.ShowConfig)
                    {
                        config.OutputConfig();
                    }
                }
                else
                {
                    if (config.ShowConfig)
                    {
                        config.OutputConfig();
                    }

                    var ctx = new AppContext
                    {
                        Config = config
                    };

                    IAppCommand command = default;

                    if (config.ListStores)
                    {
                        command = serviceProvider.GetRequiredService<ListStoresCommand>();
                    }
                    else
                    {
                        command = serviceProvider.GetRequiredService<SaveAttachmentsCommand>();
                    }
                    if (command != default)
                    {
                        exitCode = command.Run(ctx);
                    }
                }
            }

            return exitCode;
        }
    }
}
