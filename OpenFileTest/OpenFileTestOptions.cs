using CommandLine;
using CommandLine.Text;


namespace OpenFileTest
{
    public class OpenFilesTestOptions
    {
        [Option("f", Required = false, HelpText = "Source FILE to process")]
        public string file { get; set; }

        [Option("d", Required = false, HelpText = "Source DIRECTORY to search for the files specified by 'filetype'")]
        public string directory { get; set; }

        [Option("filetype", Required = false, DefaultValue = "*", HelpText = "Which type of MSOffice Documents to check (and the wildcard extension to use) doc|ppt|xls|docx|xlsx|pptx|docm|xlsm|pptm. If omitted, chose filetype based on extension.")]
        public string filetype { get; set; }

        [Option("l", Required = false, HelpText = "generate a log report of the results")]
        public bool log { get; set; }

        [Option("x", Required = false, HelpText = "Generate an XML report containing the test run results")]
        public string xmlreport { get; set; }

        [Option("o", Required = false, HelpText = "optional output directory")]
        public string outputDirectory { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo(
                    OpenFilesTestMain.GetThisApplicationName(),
                    OpenFilesTestMain.GetThisApplicationVersion()
                    ),

                Copyright = new CopyrightInfo("Glasswall", 2016),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };
            help.AddPreOptionsLine("Usage: OpenFilesTest --d Directorytoprocess [--filetype doc|ppt|xls|docx|xlsx|pptx|docm|xlsm|pptm] [other options] ");
            help.AddOptions(this);
            return help;
        }



    }
}
