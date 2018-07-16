using System;
using Microsoft.Office.Interop.Word;

using CSharpFunctionalExtensions;

namespace WordInteroper
{
    internal static class WordAppExtensions
    {
        public static Result Replace(
            this _Application app, 
            string textToFind, 
            string replacement, 
            WdReplace wordReplace,
            Action<ReplaceOptions> configureOptions = null)
        {
            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = textToFind;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replacement;

            return findObject.Replace(wordReplace, configureOptions);
        }
    }
}