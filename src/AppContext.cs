
namespace SaveOutlookAttachments
{
    using System;
    using System.IO;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class AppContext
    {
        public string StoreName { get; set; }
        public string SourcePst { get; set; }
        public string TargetFolder { get; set; }

        public Statistics Stats { get; set; } = new Statistics();
        public Outlook.NameSpace Session { get; set; }

        public static AppContext GetAppContextFromConfig(AppConfig cfg)
        {
            var result = new AppContext();

            result.StoreName = cfg.StoreName;
            result.SourcePst = GetPathWithEnvVars(cfg.SourcePst);
            result.TargetFolder = GetPathWithEnvVars(cfg.TargetFolder);

            return result;
        }

        private static string GetPathWithEnvVars(string s)
        {
            string result = Environment.ExpandEnvironmentVariables(s);
            return Path.GetFullPath(result);
        }
    }
}
