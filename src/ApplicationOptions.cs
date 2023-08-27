using System;

namespace SaveOutlookAttachments
{
    public class ApplicationOptions
    {
        public bool ShowHelp { get; internal set; }
        public bool ShowConfig { get; internal set; }
        public bool ListStores { get; internal set; }
        public string StoreDescriptor { get; internal set; }
        public string TargetFolder { get; internal set; }
        public bool WhatIf { get; internal set; }

        public ApplicationOptions()
        {
            TargetFolder = @".\Attachments";
        }

        public ApplicationOptions UseCommandLine(string[] args)
        {
            var state = ParseState.ExpectOption;

            foreach (var arg in args)
            {
                switch (state)
                {
                    case ParseState.ExpectOption:
                        state = ParseCommandLineArg(arg);
                        break;
                    case ParseState.ExpectStore:
                        StoreDescriptor = arg;
                        state = ParseState.ExpectOption;
                        break;
                    case ParseState.ExpectTargetFolder:
                        TargetFolder = FileHelpers.GetPathWithEnvVars(arg);
                        state = ParseState.ExpectOption;
                        break;
                    default:
                        break;
                }
            }

            return this;
        }

        public static void OutputHelp()
        {
            Console.WriteLine("Save Outlook Attachments.");
            Console.WriteLine();
            Console.WriteLine("SaveOutlookAttachments [-?] -s name [-p path] [-t path]");
            Console.WriteLine("  -?                 Display this help information.");
            Console.WriteLine("  -c                 Display the current configuration information.");
            Console.WriteLine("  -l                 List the stores in the default MAPI profile.");
            Console.WriteLine("  -s name or path    Store name of file path to a PST file.");
            Console.WriteLine("  -t path            Target folder for the attachments (default is \".\\Attachments\").");
            Console.WriteLine();
        }

        public void OutputOptions()
        {
            Console.WriteLine($"ListStores=\"{ListStores}\"");
            Console.WriteLine($"StoreDescriptor=\"{StoreDescriptor}\"");
            Console.WriteLine($"TargetFolder=\"{TargetFolder}\"");
            Console.WriteLine($"WhatIf={WhatIf}");
            Console.WriteLine();
        }

        private enum ParseState
        {
            ExpectOption,
            ExpectStore,
            ExpectTargetFolder
        }

        private ParseState ParseCommandLineArg(string arg)
        {
            ParseState result = ParseState.ExpectOption;

            if (arg.Length > 1 && (arg.StartsWith("-") || arg.StartsWith("/")))
            {
                var opt = arg.TrimStart('-', '/').ToLowerInvariant();
                switch (opt)
                {
                    case "?":
                        ShowHelp = true;
                        break;
                    case "c":
                        ShowConfig = true;
                        break;
                    case "l":
                        ListStores = true;
                        break;
                    case "s":
                        result = ParseState.ExpectStore;
                        break;
                    case "t":
                        result = ParseState.ExpectTargetFolder;
                        break;
                    case "whatif":
                        WhatIf = true;
                        break;
                    default:
                        break;
                }
            }
            return result;
        }
    }
}
