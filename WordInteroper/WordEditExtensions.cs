using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using CSharpFunctionalExtensions;
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

        public static Result ReplaceTokens(this WordEdit word, IReadOnlyList<FindReplace> tokenReplacements)
        {
            Contracts.Require(tokenReplacements.Any(), "No tokens provided.");
            
            foreach (FindReplace item in tokenReplacements)
            {
                Result replaceResult = word.Application.Replace(item.Token, item.Replacement, Word.WdReplace.wdReplaceAll);

                if (replaceResult.IsFailure)
                {
                    return replaceResult;
                }
            }

            return Result.Ok();
        }
    }
}
