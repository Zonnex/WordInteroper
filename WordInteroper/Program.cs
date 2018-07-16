using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using static System.Console;
using Word = Microsoft.Office.Interop.Word;
using CSharpFunctionalExtensions;
using System.Collections.Generic;

namespace WordInteroper
{
    class Program
    {
        private static readonly List<FindReplace> TokenReplacements = new List<FindReplace>
        {
            new FindReplace { Token = "__Text to replace__", Replacement = "TokenReplacement" },
            new FindReplace { Token = "__Token2__", Replacement = "Token2Replacement" },
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
                    WriteLine(replaceTokensResult.Error);
                    WriteLine("Press any key to terminate the app. No changes will be made.");
                    ReadKey();
                    return;
                }

                string pdfFilePath = Path.ChangeExtension(wordEdit.OriginalFile.FullName, "pdf");
                Result saveAsPdfResult = wordEdit.ExportAsPdf(pdfFilePath);

                string outputMessage = saveAsPdfResult.IsSuccess
                    ? $"Pdf saved at: {pdfFilePath}"
                    : saveAsPdfResult.Error;

                WriteLine(outputMessage);
            }

            WriteLine("All done. Press any key to quit");
        }

        public static void EnsureProcessClosed(string processName)
        {
            Contracts.Require(processName.HasValue());

            Maybe<Process> processResult = Process.GetProcesses()
                .FirstOrDefault(p => p.ProcessName.Contains(processName, StringComparison.OrdinalIgnoreCase));

            if (processResult.HasValue)
            {
                WriteLine("Detected Word process active. Terminating process.");
                processResult.Value.Kill();
            }
        }

        public static string GetUserInput(string message)
        {
            WriteLine(message);
            return ReadLine();
        }
    }
}