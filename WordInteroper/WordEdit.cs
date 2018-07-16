using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CSharpFunctionalExtensions;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    public class WordEdit : IDisposable
    {
        public static WordEdit OpenEdit(FileInfo wordFile)
        {
            Contracts.Require(wordFile.Extension == ".docx", "only .docx files supported");

            var app = new Word.Application();
            Word.Document document = app.Documents.Open(wordFile.FullName, ReadOnly: false, Visible: true);
            return new WordEdit
            {
                Application = app,
                Document = document,
                OriginalFile = wordFile
            };
        }
        public Word.Application Application { get; private set; }
        public Word.Document Document  { get; private set; }
        public FileInfo OriginalFile { get; private set; }

        public Result ExportAsPdf(string path)
        {
            return Document.ExportAsPdf(path);
        }

        public Result ReplaceTokens(IReadOnlyList<FindReplace> tokenReplacements)
        {
            Contracts.Require(tokenReplacements.Any(), "No tokens provided.");

            foreach (FindReplace item in tokenReplacements)
            {
                Result replaceResult = Application.Application.Replace(item.Token, item.Replacement, Word.WdReplace.wdReplaceAll);

                if (replaceResult.IsFailure)
                {
                    return replaceResult;
                }
            }

            return Result.Ok();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (Document != null)
                {
                    Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(Document);
                }

                if (Application != null)
                {
                    Application.Quit();
                    Marshal.ReleaseComObject(Application);
                    Marshal.FinalReleaseComObject(Application);
                }
            }
        }
    }
}
