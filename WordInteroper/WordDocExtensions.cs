using System;
using System.IO;
using CSharpFunctionalExtensions;

namespace WordInteroper
{
    public static class WordDocumentExtensions
    {
        public static Result ExportAsPdf(this Microsoft.Office.Interop.Word._Document doc, string filePath)
        {
            try
            {
                doc.ExportAsFixedFormat(filePath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                return Result.Ok();
            }
            catch(Exception ex)
            {
                return Result.Fail(ex.Message);
            }
        }
    }
}