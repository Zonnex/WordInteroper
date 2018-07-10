using System;
using System.IO;
using CSharpFunctionalExtensions;

namespace WordInteroper
{
    public static class WordDocExtensions
    {
        public static Result ExportAsPdf(this Microsoft.Office.Interop.Word._Document doc, FileInfo file)
        {
            try
            {
                doc.ExportAsFixedFormat(file.FullName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                return Result.Ok();
            }
            catch(Exception ex)
            {
                return Result.Fail(ex.Message);
            }
        }
    }
}