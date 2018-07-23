using System;
using System.Collections.Generic;
using System.Linq;
using CommandLine;
using SelectQAForMe.Tools;

namespace SelectQAForMe
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Starting!");
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(RunOptionsAndReturnExitCode);
        }

        private static void RunOptionsAndReturnExitCode(Options opts)
        {
            var fileParser = new QAFileParser(opts);
            var result = fileParser.ExtractDataToOutFile();
        }
    }
}