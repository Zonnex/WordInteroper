using System;
using CSharpFunctionalExtensions;
using Microsoft.Office.Interop.Word;
using WordInteroper.Models;

namespace WordInteroper.Extensions
{
    internal static class WordAppExtensions
    {
        public static Result ReplaceToken(
            this _Application app, 
            TokenReplacement tokenReplace, 
            WdReplace wordReplace = WdReplace.wdReplaceAll,
            Action<ReplaceOptions> configureOptions = null)
        {
            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = tokenReplace.Token;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = tokenReplace.Replacement;

            return findObject.Replace(wordReplace, configureOptions);
        }
    }
}