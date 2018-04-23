using System;
using System.Drawing;
using DevExpress.Spreadsheet;

namespace SpreadsheetControl_WPF_API
{
    public static class CellActions {
        #region Actions
        public static Action<IWorkbook> ChangeCellValueAction = ChangeCellValue;
        public static Action<IWorkbook> AddHyperlinkAction = AddHyperlink;
        public static Action<IWorkbook> CopyCellDataAndStyleAction = CopyCellDataAndStyle;
        public static Action<IWorkbook> MergeAndSplitCellsAction = MergeAndSplitCells;
        public static Action<IWorkbook> ClearCellsAction = ClearCells;
        #endregion

        static void ChangeCellValue(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["A1"].Value = "dateTime:";
                worksheet.Cells["A2"].Value = "double:";
                worksheet.Cells["A3"].Value = "string:";
                worksheet.Cells["A4"].Value = "error constant:";
                worksheet.Cells["A5"].Value = "boolean:";
                worksheet.Cells["A6"].Value = "float:";
                worksheet.Cells["A7"].Value = "char:";
                worksheet.Cells["A8"].Value = "int32:";
                worksheet.Cells["A10"].Value = "Fill a range of cells:";
                worksheet.Columns["A"].WidthInCharacters = 20;
                worksheet.Columns["B"].WidthInCharacters = 20;
                worksheet.Range["A1:B8"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;

                #region #CellValue
                // Add data of different types to cells.
                worksheet.Cells["B1"].Value = DateTime.Now;
                worksheet.Cells["B2"].Value = Math.PI;
                worksheet.Cells["B3"].Value = "Have a nice day!";
                worksheet.Cells["B4"].Value = CellValue.ErrorReference;
                worksheet.Cells["B5"].Value = true;
                worksheet.Cells["B6"].Value = float.MaxValue;
                worksheet.Cells["B7"].Value = 'a';
                worksheet.Cells["B8"].Value = Int32.MaxValue;

                // Fill all cells in the range with 10.
                worksheet.Range["B10:E10"].Value = 10;
                #endregion #CellValue
            }
            finally {
                workbook.EndUpdate();
            }

        }

        static void AddHyperlink(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Range["A:C"].ColumnWidthInCharacters = 12;

                #region #AddHyperlink
                // Create a hyperlink to a web page.
                Cell cell = worksheet.Cells["A1"];
                worksheet.Hyperlinks.Add(cell, "http://www.devexpress.com/", true, "DevExpress");

                // Create a hyperlink to a cell range in a workbook.
                Range range = worksheet.Range["C3:D4"];
                Hyperlink cellHyperlink = worksheet.Hyperlinks.Add(range, "Sheet2!B2:E7", false, "Select Range");
                cellHyperlink.TooltipText = "Click Me";
                #endregion #AddHyperlink
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void CopyCellDataAndStyle(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                #region #CopyCell
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Columns["A"].WidthInCharacters = 32;
                worksheet.Columns["B"].WidthInCharacters = 20;
                Style style = workbook.Styles[BuiltInStyleId.Input];

                // Specify the content and formatting for a source cell.
                worksheet.Cells["A1"].Value = "Source Cell";

                Cell sourceCell = worksheet.Cells["B1"];
                sourceCell.Formula = "= PI()";
                sourceCell.NumberFormat = "0.0000";
                sourceCell.Style = style;
                sourceCell.Font.Color = Color.Blue;
                sourceCell.Font.Bold = true;
                sourceCell.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

                // Copy all information from the source cell to the "B3" cell. 
                worksheet.Cells["A3"].Value = "Copy All";
                worksheet.Cells["B3"].CopyFrom(sourceCell);

                // Copy only the source cell content (e.g., text, numbers, formula calculated values) to the "B4" cell.
                worksheet.Cells["A4"].Value = "Copy Values";
                worksheet.Cells["B4"].CopyFrom(sourceCell, PasteSpecial.Values);

                // Copy the source cell content (e.g., text, numbers, formula calculated values) 
                // and number formats to the "B5" cell.
                worksheet.Cells["A5"].Value = "Copy Values and Number Formats";
                worksheet.Cells["B5"].CopyFrom(sourceCell, PasteSpecial.Values | PasteSpecial.NumberFormats);

                // Copy only the formatting information from the source cell to the "B6" cell.
                worksheet.Cells["A6"].Value = "Copy Formats";
                worksheet.Cells["B6"].CopyFrom(sourceCell, PasteSpecial.Formats);

                // Copy all information from the source cell to the "B7" cell except for border settings.
                worksheet.Cells["A7"].Value = "Copy All Except Borders";
                worksheet.Cells["B7"].CopyFrom(sourceCell, PasteSpecial.All & ~PasteSpecial.Borders);

                // Copy information only about borders from the source cell to the "B8" cell.
                worksheet.Cells["A8"].Value = "Copy Borders";
                worksheet.Cells["B8"].CopyFrom(sourceCell, PasteSpecial.Borders);
                #endregion #CopyCell
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void MergeAndSplitCells(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["A2"].FillColor = Color.LightGray;

                worksheet.Cells["B2"].Value = "B2";
                worksheet.Cells["B2"].FillColor = Color.LightGreen;

                worksheet.Cells["C3"].Value = "C3";
                worksheet.Cells["C3"].FillColor = Color.LightSalmon;

                #region #MergeCells
                // Merge cells contained in the range.
                worksheet.MergeCells(worksheet.Range["A1:C5"]);
                // Split merged cells contained in the range.
                worksheet.UnMergeCells(worksheet.Range["A1:C5"]);
                // Merge cells contained in the range.
                worksheet.MergeCells(worksheet.Range["A1:C5"]);
                #endregion #MergeCells
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void ClearCells(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A:D"].ColumnWidthInCharacters = 30;
                worksheet.Range["B1:D6"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                worksheet["B1"].Value = "Initial Cell Content and Formatting:";
                worksheet.MergeCells(worksheet["C1:D1"]);
                worksheet["C1:D1"].Value = "Cleared Cells:";

                worksheet["A2"].Value = "Clear All:";
                worksheet["A3"].Value = "Clear Cell Content Only:";
                worksheet["A4"].Value = "Clear Cell Formatting Only:";
                worksheet["A5"].Value = "Clear Cell Hyperlinks Only:";

                // Specify initial content and formatting for cells.
                Range sourceCells = worksheet["B2:D5"];
                sourceCells.Value = DateTime.Now;
                sourceCells.Style = workbook.Styles[BuiltInStyleId.Accent3_40percent];
                sourceCells.Font.Color = Color.LightSeaGreen;
                sourceCells.Font.Bold = true;
                sourceCells.Borders.SetAllBorders(Color.Blue, BorderLineStyle.Dashed);
                worksheet.Hyperlinks.Add(worksheet["B5"], "http://www.devexpress.com/", true, "DevExpress");
                worksheet.Hyperlinks.Add(worksheet["C5"], "http://www.devexpress.com/", true, "DevExpress");
                worksheet.Hyperlinks.Add(worksheet["D5"], "http://www.devexpress.com/", true, "DevExpress");

                #region #ClearCell
                // Remove all cell information (content, formatting, hyperlinks and comments).
                worksheet.Clear(worksheet["C2:D2"]);

                // Remove cell content.
                worksheet.ClearContents(worksheet["C3"]);
                worksheet["D3"].Value = null;

                // Remove cell formatting.
                worksheet.ClearFormats(worksheet["C4"]);
                worksheet["D4"].Style = workbook.Styles.DefaultStyle;

                // Remove hyperlinks from cells.
                worksheet.ClearHyperlinks(worksheet["C5"]);

                Hyperlink hyperlinkD5 = worksheet.Hyperlinks.GetHyperlinks(worksheet["D5"])[0];
                worksheet.Hyperlinks.Remove(hyperlinkD5);
                #endregion #ClearCell
            }
            finally {
                workbook.EndUpdate();
            }
        }
    }
}
