using System;
using System.Runtime.InteropServices;
using CSharpFunctionalExtensions;

namespace WordInteroper
{
    public static class WordFindExtensions
    {
        public class ReplaceOptions
        {
            public static ReplaceOptions Default = new ReplaceOptions();
            public bool MatchCase = false;
            public bool MatchWholeWord = true;
            public bool MatchWildcards = false;
            public bool MatchSoundsLike = false;
            public bool MatchAllWordForms = false;
            public bool Forward = true;
            public bool Format = false;
            public bool MatchKashida = false;
            public bool MatchDiacritics = false;
            public bool MatchAlefHamza = false;
            public bool MatchControl = false;
            public int Wrap = 1;
        }

        public static Result Replace(
            this Microsoft.Office.Interop.Word.Find find,
            Microsoft.Office.Interop.Word.WdReplace wordReplace,
            Action<ReplaceOptions> configureOptions = null)
        {
            ReplaceOptions options = ReplaceOptions.Default;
            configureOptions?.Invoke(options);
            try
            {
                bool success = find.Execute(
                    FindText: find.Text,
                    MatchCase: options.MatchCase,
                    MatchWholeWord: options.MatchWholeWord,
                    MatchWildcards: options.MatchWildcards,
                    MatchSoundsLike: options.MatchSoundsLike,
                    MatchAllWordForms: options.MatchAllWordForms,
                    Forward: options.Forward,
                    Wrap: options.Wrap,
                    Format: options.Format,
                    ReplaceWith: find.Replacement.Text,
                    Replace: wordReplace,
                    MatchKashida: options.MatchKashida,
                    MatchDiacritics: options.MatchDiacritics,
                    MatchAlefHamza: options.MatchAlefHamza,
                    MatchControl: options.MatchControl);

                return success
                    ? Result.Ok()
                    : Result.Fail("Something went wrong");
            }
            catch (COMException ex)
            {
                return Result.Fail(ex.Message + ". Check if word is already open prior to running program");
            }
            catch (Exception ex)
            {
                return Result.Fail(ex.Message);
            }
        }
    }
}