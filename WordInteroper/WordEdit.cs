using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CSharpFunctionalExtensions;
using WordInteroper.Extensions;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    public class WordEdit : IDisposable
    {
        public static WordEdit OpenEdit(string filePath)
        {
            Contracts.Require(filePath.HasValue());
            Contracts.Require(Path.GetExtension(filePath) == ".docx", "only .docx files supported");

            return OpenEdit(new FileInfo(filePath));
        }

        public static WordEdit OpenEdit(FileInfo wordFile)
        {
            Contracts.Require(wordFile.Exists);
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
            Contracts.Require(path.HasValue());
            Contracts.Require(Path.GetExtension(path) == ".pdf");
            try
            {
                Document.ExportAsFixedFormat(path, Word.WdExportFormat.wdExportFormatPDF);
                return Result.Ok();
            }
            catch (Exception ex)
            {
                return Result.Fail(ex.Message);
            }
        }

        public Result ReplaceTokens(IReadOnlyList<TokenReplace> tokenReplacements)
        {
            Contracts.Require(tokenReplacements.Any(), "No tokens provided.");

            foreach (TokenReplace item in tokenReplacements)
            {
                Result replaceResult = Application.ReplaceToken(item);

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
