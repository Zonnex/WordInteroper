using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CSharpFunctionalExtensions;
using static System.Console;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    class Program
    {
        private const string Token = "__Text to replace__";
        private const string File1 = @"C:\Users\connys\Desktop\work\SOSAlarm\EditWordDocument\word_openxml.docx";

        private static void Main(string[] args)
        {
            Word.Application app = null;
            Word.Document document = null;
            try
            {
                Maybe<Process> processResult = CheckProcessIsRunning();
                if (processResult.HasValue)
                {
                    processResult.Value.Kill();
                }
                app = new Word.Application();
                document = app.Documents.Open(File1, ReadOnly: false, Visible: true);
                Work(app, document);
            }
            catch (Exception ex)
            {
                DebugWrite(ex.Message);
            }
            finally
            {
                ReleaseResources(app, document);
            }

            DebugWrite("All done. Press any key to quit");
        }

        private static void ReleaseResources(Word._Application app, Word._Document document)
        {
            if (document != null)
            {
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Marshal.ReleaseComObject(document);
            }

            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
                Marshal.FinalReleaseComObject(app);
            }
        }

        private static Maybe<Process> CheckProcessIsRunning()
        {
            return Process.GetProcesses()
                .FirstOrDefault(p => p.ProcessName.Contains("WINWORD", StringComparison.OrdinalIgnoreCase));
        }

        [Conditional("DEBUG")]
        private static void DebugWrite(string message)
        {
            WriteLine(message);
        }

        private static void Work(Word._Application app, Word._Document doc)
        {
            const string replacement = "some replacement";
            DebugWrite($"Find:{Token} Replace: {replacement}");
            Result replaceResult = app.Replace(Token, replacement, Word.WdReplace.wdReplaceAll);
            if (replaceResult.IsSuccess)
            {
                string path = Path.Combine(Path.GetDirectoryName(File1), "test.pdf");
                Result result = doc.ExportAsPdf(new FileInfo(path));

                if (result.IsSuccess)
                {
                    WriteLine("Replaced and saved document");
                }
            }
            else
            {
                WriteLine(replaceResult.Error);
            }
        }
    }
}
