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
                var options = new ApplicationOptions()
                    .UseCommandLine(args);

                if (options.ShowHelp)
                {
                    ApplicationOptions.OutputHelp();
                    if (options.ShowConfig)
                    {
                        options.OutputOptions();
                    }
                }
                else
                {
                    if (options.ShowConfig)
                    {
                        options.OutputOptions();
                    }

                    var ctx = new AppContext
                    {
                        Options = options
                    };

                    IAppCommand command = default;

                    if (options.ListStores)
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
