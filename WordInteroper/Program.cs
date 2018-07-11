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
    partial class Program
    {
        private const string Token = "__Text to replace__";
        private const string FilePath = @"C:\Users\connys\Desktop\work\SOSAlarm\EditWordDocument\word_openxml.docx";

        private static List<FindReplace> TokenReplacements = new List<FindReplace>
        {
            new FindReplace { Token = "__Text to replace__", Replacement = "TokenReplacement" },
        };

        private static void Main(string[] args)
        {
            WordEdit wordEdit = null;
            try
            {
                EnsureProcessClosed("WINWORD");
                FileInfo fileInfo = new FileInfo(FilePath);
                wordEdit = WordEdit.OpenEdit(fileInfo);
                Result replaceTokensResult = ReplaceTokens(wordEdit, TokenReplacements);
                
                if(replaceTokensResult.IsFailure)
                {
                    WriteLine(replaceTokensResult.Error);
                    WriteLine("Press any key to terminate the app. No changes will be made.");
                    ReadKey();
                    return;
                }

                string pdfFilePath = Path.ChangeExtension(wordEdit.OriginalFile.FullName, "pdf");
                Result saveAsPdfResult = wordEdit.Document.ExportAsPdf(pdfFilePath);

                if(saveAsPdfResult.IsSuccess)
                    WriteLine($"Pdf saved at: {pdfFilePath}");
                else
                    WriteLine(saveAsPdfResult.Error);
            }
            catch (Exception ex)
            {
                DebugWrite(ex.Message);
            }
            finally
            {
                wordEdit.ReleaseResources();
            }

            DebugWrite("All done. Press any key to quit");
        }

        public static void EnsureProcessClosed(string processName)
        {
            Contracts.Require(processName.HasValue());

            Maybe<Process> processResult = Process.GetProcesses()
                .FirstOrDefault(p => p.ProcessName.Contains(processName, StringComparison.OrdinalIgnoreCase));

            if (processResult.HasValue)
            {
                processResult.Value.Kill();
            }
        }

        [Conditional("DEBUG")]
        public static void DebugWrite(string message)
        {
            WriteLine(message);
        }

        public static Result ReplaceTokens(WordEdit word, IReadOnlyList<FindReplace> tokenReplacements)
        {
            Contracts.Require(tokenReplacements.Any(), "No tokens provided.");
            foreach (FindReplace item in tokenReplacements)
            {
                Result replaceResult = word.Application.Replace(item.Token, item.Replacement, Word.WdReplace.wdReplaceAll);

                if (replaceResult.IsFailure)
                {
                    WriteLine(replaceResult.Error);
                    return replaceResult;
                }
            }

            return Result.Ok();
        }
    }
}
