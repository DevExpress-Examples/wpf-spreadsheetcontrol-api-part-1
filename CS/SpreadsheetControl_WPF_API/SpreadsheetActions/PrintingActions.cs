using System;
using System.Drawing;
#region #printingUsings
using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting;
using DevExpress.Xpf.Printing;
using System.Windows.Controls; 
#endregion #printingUsings

namespace SpreadsheetControl_WPF_API
{
    public static class PrintingActions {

        public static Action<IWorkbook> PrintAction = Print;

        static void Print(IWorkbook workbook) {

            Worksheet worksheet = workbook.Worksheets[0];

            // Generate worksheet content - the simple multiplication table.
            Range columnHeadings = worksheet.Range.FromLTRB(1,0,40,0);
            columnHeadings.Formula = "=COLUMN() - 1";
            Range rowHeadings = worksheet.Range.FromLTRB(0,1,0,40);
            rowHeadings.Formula = "=ROW() - 1";
            Range tableRange = worksheet.Range.FromLTRB(1,1,40,40);
            tableRange.Formula = "=(ROW()-1)*(COLUMN()-1)";

            // Format headers of the multiplication table.
            Formatting rangeFormatting = columnHeadings.BeginUpdateFormatting();
            rangeFormatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin;
            rangeFormatting.Borders.BottomBorder.Color = Color.Black;
            columnHeadings.EndUpdateFormatting(rangeFormatting);

            rangeFormatting = rowHeadings.BeginUpdateFormatting();
            rangeFormatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin;
            rangeFormatting.Borders.RightBorder.Color = Color.Black;
            rowHeadings.EndUpdateFormatting(rangeFormatting);

            rangeFormatting = tableRange.BeginUpdateFormatting();
            rangeFormatting.Fill.BackgroundColor = Color.LightBlue;
            tableRange.EndUpdateFormatting(rangeFormatting);

            #region #WorksheetPrintOptions
            worksheet.ActiveView.Orientation = PageOrientation.Landscape;
            //  Display row and column headings.
            worksheet.ActiveView.ShowHeadings = true;
            worksheet.ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4;
            // Access an object that contains print options.
            WorksheetPrintOptions printOptions = worksheet.PrintOptions;
            //  Print in black and white.
            printOptions.BlackAndWhite = true;
            //  Do not print gridlines.
            printOptions.PrintGridlines = false;
            //  Scale the print area to fit to two pages wide.
            printOptions.FitToPage = true;
            printOptions.FitToWidth = 2;
            //  Print a dash instead of a cell error message.
            printOptions.ErrorsPrintMode = ErrorsPrintMode.Dash;
            #endregion #WorksheetPrintOptions

            #region #PrintWorkbook
            // Invoke the Print Preview dialog for the workbook.
                using (LegacyPrintableComponentLink link = new LegacyPrintableComponentLink(workbook))
                {
                    link.CreateDocument();
                    link.ShowPrintPreviewDialog(App.Current.MainWindow);
                }
            #endregion #PrintWorkbook
        }
    }
}
