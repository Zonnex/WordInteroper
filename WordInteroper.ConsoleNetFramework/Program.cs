using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using CSharpFunctionalExtensions;
using WordInteroper.Extensions;

namespace WordInteroper.ConsoleNetFramework
{
    class Program
    {
        private static readonly List<TokenReplacement> TokenReplacements = new List<TokenReplacement>
        {
            new TokenReplacement { Token = "__DATE__", Replacement = DateTime.Today.ToString(CultureInfo.InvariantCulture)},
        };

        private static void Main(string[] args)
        {
            EnsureProcessClosed("WINWORD");
            //var fileInfo = new FileInfo(GetUserInput("Enter file path to docx file"));
            var fileInfo = new FileInfo(@"C:\dev\ActiveSolution\SOSAlarm\OrderWeb\OrderWeb.WordInterop\Document\Transportmeddelande vid transport mellan vårdenheter 2.0.docx");
            using (WordEdit wordEdit = WordEdit.OpenEdit(fileInfo))
            {
                Result replaceTokensResult = wordEdit.ReplaceTokens(TokenReplacements);

                if (replaceTokensResult.IsFailure)
                {
                    Console.WriteLine(replaceTokensResult.Error);
                    Console.WriteLine("Press any key to close the program. No changes will be made.");
                    Console.ReadKey();
                    return;
                }
                
                FileInfo outputFile = wordEdit.File.ChangeExtension("pdf");
                Result saveAsPdfResult = wordEdit.ExportAsPdf(outputFile.FullName);

                string outputMessage = saveAsPdfResult.IsSuccess
                    ? $"Pdf saved at: {outputFile.FullName}"
                    : saveAsPdfResult.Error;

                Console.WriteLine(outputMessage);
            }

            Console.WriteLine("All done. Press any key to quit");
        }

        public static void EnsureProcessClosed(string processName)
        {
            Contracts.Require(processName.HasValue());

            Maybe<Process> processResult = Process.GetProcesses()
                .FirstOrDefault(p => p.ProcessName.Contains(processName, StringComparison.OrdinalIgnoreCase));

            if (processResult.HasValue)
            {
                Console.WriteLine("Detected Word process active. Terminating process.");
                processResult.Value.Kill();
            }
        }

        public static string GetUserInput(string message)
        {
            Console.WriteLine(message);
            return Console.ReadLine();
        }
    }
}