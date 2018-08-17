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
    public interface IWordCheckBox
    {
        string Title { get; }
        string Tag { get; }
        bool Checked { get; }
    }

    public enum LoadMode
    {
        FullDispose,
        KeepAlive,
    }

    public class WordEdit : IDisposable
    {
        public static WordEdit OpenEdit(string filePath)
        {
            Contracts.Require(filePath.HasValue());

            return OpenEdit(new FileInfo(filePath));
        }

        public static WordEdit OpenEdit(string filePath, LoadMode loadMode)
        {
            Contracts.Require(filePath.HasValue());

            return OpenEdit(new FileInfo(filePath), loadMode);
        }

        public static WordEdit OpenEdit(FileInfo wordFile)
        {
            return OpenEdit(wordFile, LoadMode.FullDispose);
        }

        public static WordEdit OpenEdit(FileInfo wordFile, LoadMode loadMode)
        {
            Contracts.Require(wordFile.Exists);
            Contracts.Require(wordFile.Extension == ".docx", "only .docx files supported");

            Word.Application app = OpenWord();
            Word.Document document = app.Documents.Open(wordFile.FullName, ReadOnly: false, Visible: true);

            return new WordEdit
            {
                Application = app,
                Document = document,
                File = wordFile,
                LoadMode = loadMode
            };

            Word.Application OpenWord()
            {
                return loadMode == LoadMode.FullDispose
                    ? new Word.Application()
                    : Marshal.GetActiveObject("Word.Application") as Word.Application;
            }
        }

        public Word.Application Application { get; private set; }
        public Word.Document Document  { get; private set; }
        public FileInfo File { get; private set; }
        protected LoadMode LoadMode { get; private set; }

        public Result SetCheckboxes(IEnumerable<IWordCheckBox> checkboxValues)
        {
            // ReSharper disable once PossibleMultipleEnumeration
            Contracts.RequireNotNull(checkboxValues);

            (dynamic WordCheckbox, IWordCheckBox SetCheckBox)[] valueTuples = Document.GetCheckboxes()
                .Join(checkboxValues, d => (d.Title, d.Tag), c => (c.Title, c.Tag), (d, c) => (WordCheckbox: d, SetCheckBox: c))
                .ToArray();

            foreach ((dynamic wordCheckbox, IWordCheckBox wordCheckBox) in valueTuples)
            {
                wordCheckbox.Checked = wordCheckBox.Checked;
            }
            
            return Result.Ok();
        }

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

        public Result ReplaceTokens(IReadOnlyList<TokenReplacement> tokenReplacements)
        {
            Contracts.RequireNotNull(tokenReplacements);

            foreach (TokenReplacement item in tokenReplacements)
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

                if (Application != null && LoadMode == LoadMode.FullDispose)
                {
                    Application.Quit();
                    Marshal.ReleaseComObject(Application);
                    Marshal.FinalReleaseComObject(Application);
                }
            }
        }
    }
}
