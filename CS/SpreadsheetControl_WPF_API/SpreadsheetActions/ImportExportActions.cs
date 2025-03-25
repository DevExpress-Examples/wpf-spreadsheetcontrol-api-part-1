using System;
using System.IO;
using System.Diagnostics;
using DevExpress.Spreadsheet;

namespace SpreadsheetControl_WPF_API
{
    public static class ImportExportActions {

        public static Action<IWorkbook> ExportToPdfAction = ExportToPdf;

        static void ExportToPdf(IWorkbook workbook) {
            workbook.Worksheets[0].Cells["D8"].Value = "This document is exported to the PDF format.";

            #region #ExportToPdf
            using (FileStream pdfFileStream = new FileStream("Documents\\Document_PDF.pdf", FileMode.Create)) {
                workbook.ExportToPdf(pdfFileStream);
            }
            #endregion #ExportToPdf
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"Documents\Document_PDF.pdf")
            {
                UseShellExecute = true
            };
            p.Start();
        }
    }
}
