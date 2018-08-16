using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordInteroper.Extensions
{
    internal static class WordDocumentExtensions
    {
        [DebuggerStepThrough]
        public static IEnumerable<dynamic> GetCheckboxes(this Document document)
        {
            return document.GetObjects()
                .Where(d => d.Type == 8);
        }

        [DebuggerStepThrough]
        public static IEnumerable<dynamic> GetObjects(this Document document)
        {
            foreach (dynamic contentControl in document.ContentControls)
                yield return contentControl;
        }
    }
}