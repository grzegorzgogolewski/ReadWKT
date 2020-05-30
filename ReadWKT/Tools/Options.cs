using CommandLine;

namespace ReadWKT.Tools
{
    public class Options
    {
        [Option('s', "startupPath", Required = true, HelpText = "Katalog z danymi, np. -s c:\\temp")]
        public string StarupPath { get; set; }
    }
}