using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using CSharpFunctionalExtensions;

namespace WordInteroper.ConsoleNetFramework
{
    class Program
    {
        private static readonly List<TokenReplace> TokenReplacements = new List<TokenReplace>
        {
            new TokenReplace { Token = "__Text to replace__", Replacement = "TokenReplacement" },
            new TokenReplace { Token = "__Token2__", Replacement = "Token2Replacement" },
        };

        private static void Main(string[] args)
        {
            EnsureProcessClosed("WINWORD");
            var fileInfo = new FileInfo(GetUserInput("Enter file path to docx file"));
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

                string pdfFilePath = Path.ChangeExtension(wordEdit.OriginalFile.FullName, "pdf");
                Result saveAsPdfResult = wordEdit.ExportAsPdf(pdfFilePath);

                string outputMessage = saveAsPdfResult.IsSuccess
                    ? $"Pdf saved at: {pdfFilePath}"
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