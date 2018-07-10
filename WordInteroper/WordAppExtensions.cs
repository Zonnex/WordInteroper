using CSharpFunctionalExtensions;

namespace WordInteroper
{
    public static class WordAppExtensions
    {
        public static Result Replace(
            this Microsoft.Office.Interop.Word._Application app, 
            string textToFind, 
            string replacement, 
            Microsoft.Office.Interop.Word.WdReplace wordReplace)
        {
            Microsoft.Office.Interop.Word.Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = textToFind;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replacement;

            return findObject.Replace(wordReplace);
        }
    }
}