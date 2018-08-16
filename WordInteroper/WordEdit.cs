using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CSharpFunctionalExtensions;
using Microsoft.Office.Interop.Word;
using WordInteroper.Extensions;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    public interface IWordCheckBox
    {
        string Title { get; }
        string Tag { get; }
        bool Checked { get; set; }
    }

    public class WordEdit : IDisposable
    {
        public static WordEdit OpenEdit(string filePath)
        {
            Contracts.Require(filePath.HasValue());

            return OpenEdit(new FileInfo(filePath));
        }

        public static WordEdit OpenEdit(FileInfo wordFile)
        {
            Contracts.Require(wordFile.Exists);
            Contracts.Require(wordFile.Extension == ".docx", "only .docx files supported");

            Word.Application app = new Word.Application();
            Word.Document document = app.Documents.Open(wordFile.FullName, ReadOnly: false, Visible: true);

            //Word.ContentControl contentControl = app.Selection.ContentControls.Add(Word.WdContentControlType.wdContentControlCheckBox);
            var boolValues = new []
            {
                new SetCheckBox
                {
                    Title = "Kognitiv-Svikt-Förvirring",
                    Tag = "False",
                    Checked = true
                },
                new SetCheckBox
                {
                    Title = "Kognitiv-Svikt-Förvirring",
                    Tag = "True",
                    Checked = false
                },
            };
            SetCheckboxes(document, boolValues);
            return new WordEdit
            {
                Application = app,
                Document = document,
                File = wordFile
            };

            //Word.Application OpenWord()
            //{
            //    return Marshal.GetActiveObject("Word.Application") as Word.Application
            //        ?? new Word.Application();
            //}
        }

        public Word.Application Application { get; private set; }
        public Word.Document Document  { get; private set; }
        public FileInfo File { get; private set; }

        private static void SetCheckboxes(Word.Document document, IEnumerable<IWordCheckBox> checkboxValues)
        {
            (dynamic WordCheckbox, IWordCheckBox SetCheckBox)[] valueTuples = document.GetCheckboxes()
                .Join(checkboxValues, d => (d.Title, d.Tag), c => (c.Title, c.Tag), (d, c) => (WordCheckbox: d, SetCheckBox: c))
                .ToArray();

            foreach ((dynamic WordCheckbox, IWordCheckBox SetCheckBox) valueTuple in valueTuples)
            {
                valueTuple.WordCheckbox.Checked = valueTuple.SetCheckBox.Checked;
            }
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
            Contracts.Require(tokenReplacements.Any(), "No tokens provided.");

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
