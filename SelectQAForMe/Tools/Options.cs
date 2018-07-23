using System.Runtime.InteropServices.ComTypes;
using CommandLine;
using NPOI.SS.Formula.Functions;

namespace SelectQAForMe.Tools
{
    public class Options
    {
        public Options()
        {
            WorkSheet = "Questions";
            NumberOfQuestions = 60;
        }

        [Option('f', "file", Required = true, HelpText = "Input excel file to be processed")]
        public string FileName { get; set; }

        [Option('t', "tab", Required = false, HelpText = "Worksheet to use(Default - Questions)")]
        public string WorkSheet { get; set; }

        [Option('o', "out", Required = true, HelpText = "Output file name")]
        public string OutFileName { get; set; }

        [Option('n', "number", Required = false, HelpText = "Number of Questions(default 60)")]
        public int NumberOfQuestions { get; set; }
    }
}