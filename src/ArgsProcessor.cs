
namespace SaveOutlookAttachments
{
    using System;

    public class ArgsProcessor
    {
        private enum ParseState
        {
            ExpectOption,
            ExpectStoreName,
            ExpectSourcePst,
            ExpectTargetFolder
        }
        public string SourcePst { get; set; }
        public string StoreName { get; set; }
        public string TargetFolder { get; set; }

        public static void ShowHelp()
        {
            Console.WriteLine("Save Outlook Attachments.");
            Console.WriteLine();
            Console.WriteLine("SaveOutlookAttachments [-?] -s name [-p path] [-t path]");
            Console.WriteLine("  -?                 Display this help information.");
            Console.WriteLine("  -s name            Store name.");
            Console.WriteLine("  -p path            PST file to scan for attachments.");
            Console.WriteLine("  -t path            Target folder for the attachments (default is \".\\Attachments\").");
            Console.WriteLine();
        }

        public static AppConfig Parse(string[] args)
        {
            var result = GetDefaultConfig();

            ParseState state = ParseState.ExpectOption;
            foreach (var arg in args)
            {
                switch (state)
                {
                    case ParseState.ExpectOption:
                        state = ParseOption(arg, result);
                        break;
                    case ParseState.ExpectStoreName:
                        result.StoreName = arg;
                        state = ParseState.ExpectOption;
                        break;
                    case ParseState.ExpectSourcePst:
                        result.SourcePst = arg;
                        state = ParseState.ExpectOption;
                        break;
                    case ParseState.ExpectTargetFolder:
                        result.TargetFolder = arg;
                        state = ParseState.ExpectOption;
                        break;
                    default:
                        break;
                }
            }
            return result;
        }

        private static AppConfig GetDefaultConfig()
        {
            var result = new AppConfig()
            {
                TargetFolder = @".\Attachments",
            };
            return result;
        }

        private static ParseState ParseOption(string arg, AppConfig config)
        {
            ParseState result = ParseState.ExpectOption;

            if (arg.Length > 1 && (arg[0] == '-' || arg[0] == '/'))
            {
                switch (Char.ToLowerInvariant(arg[1]))
                {
                    case '?':
                        config.ShowHelp = true;
                        break;
                    case 's':
                        result = ParseState.ExpectStoreName;
                        break;
                    case 'p':
                        result = ParseState.ExpectSourcePst;
                        break;
                    case 't':
                        result = ParseState.ExpectTargetFolder;
                        break;
                    default:
                        break;
                }
            }
            return result;
        }
    }
}
