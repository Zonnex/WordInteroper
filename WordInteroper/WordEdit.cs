using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    public class WordEdit
    {
        public static WordEdit OpenEdit(FileInfo wordFile)
        {
            Contracts.Require(wordFile.Extension == ".docx", "only .docx files supported");

            var (app, doc) = OpenDocument(wordFile);
            return new WordEdit
            {
                Application = app,
                Document = doc,
                OriginalFile = wordFile
            };
        }
        public Word.Application Application { get; private set; }
        public Word.Document Document  { get; private set; }
        public FileInfo OriginalFile { get; private set; }

        private static (Word.Application app, Word.Document document) OpenDocument(FileInfo file)
        {
            var app = new Word.Application();
            var document = app.Documents.Open(file.FullName, ReadOnly: false, Visible: true);
            return (app, document);
        }
    }
}
