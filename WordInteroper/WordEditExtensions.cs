using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordInteroper
{
    public static class WordEditExtensions
    {
        public static void ReleaseResources(this WordEdit word)
        {
            Word.Document document = word.Document;
            Word.Application app = word.Application;

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
    }
}
