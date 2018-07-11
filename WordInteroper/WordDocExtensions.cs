using System;
using System.IO;
using CSharpFunctionalExtensions;

namespace WordInteroper
{
    public static class WordDocExtensions
    {
        public static Result ExportAsPdf(this Microsoft.Office.Interop.Word._Document doc, FileInfo fileInfo)
        {
            return doc.ExportAsPdf(fileInfo);
        }

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