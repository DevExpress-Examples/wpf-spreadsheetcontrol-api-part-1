﻿using System;
using DevExpress.Spreadsheet;
using System.Drawing;

namespace SpreadsheetControl_WPF_API
{
    public static class RowAndColumnActions {
        #region Actions
        public static Action<IWorkbook> InsertRowsColumnsAction = InsertRowsColumns;
        public static Action<IWorkbook> DeleteRowsColumnsAction = DeleteRowsColumns;
        public static Action<IWorkbook> CopyRowsColumnsAction = CopyRowsColumns;
        public static Action<IWorkbook> ShowHideRowsColumnsAction = ShowHideRowsColumns;
        public static Action<IWorkbook> SpecifyRowsHeightColumnsWidthAction = SpecifyRowsHeightColumnsWidth;
        #endregion

        static void InsertRowsColumns(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate cells with data.
                for (int i = 0; i < 10; i++) {
                    worksheet.Cells[i, 0].Value = i + 1;
                    worksheet.Cells[0, i].Value = i + 1;
                }

                #region #InsertRows
                // Insert a new row 3.
                worksheet.Rows["3"].Insert();

                // Insert a new row into the worksheet at the 5the position.
                worksheet.Rows.Insert(4);

                // Insert five rows into the worksheet at the 9th position.
                worksheet.Rows.Insert(8, 5);

                // Insert two rows above the "L15:M16" cell range.
                worksheet.InsertCells(worksheet.Range["L15:M16"], InsertCellsMode.EntireRow);
                #endregion #InsertRows

                #region #InsertColumns
                // Insert a new column C.
                worksheet.Columns["C"].Insert();

                // Insert a new column into the worksheet at the 5th position.
                worksheet.Columns.Insert(4);

                // Insert three columns into the worksheet at the 7th position.
                worksheet.Columns.Insert(6, 3);

                // Insert two columns to the left of the "L15:M16" cell range.
                worksheet.InsertCells(worksheet.Range["L15:M16"], InsertCellsMode.EntireColumn);
                #endregion #InsertColumns
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void DeleteRowsColumns(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets["Sheet1"];

                // Fill cells with data.
                for (int i = 0; i < 15; i++) {
                    worksheet.Cells[i, 0].Value = i + 1;
                    worksheet.Cells[0, i].Value = i + 1;
                }

                #region #DeleteRows
                // Delete the 2nd row from the worksheet.
                worksheet.Rows[1].Delete();

                // Delete the 3rd row from the worksheet.
                worksheet.Rows.Remove(2);

                // Delete three rows from the worksheet starting from the 10th row.
                worksheet.Rows.Remove(9, 3);

                // Delete a row that contains the "B2"cell.
                worksheet.DeleteCells(worksheet.Cells["B2"], DeleteMode.EntireRow);
                #endregion #DeleteRows

                #region #DeleteColumns
                // Delete the 2nd column from the worksheet.
                worksheet.Columns[1].Delete();

                // Delete the 3rd column from the worksheet.
                worksheet.Columns.Remove(2);

                // Delete three columns from the worksheet starting from the 10th column.
                worksheet.Columns.Remove(9, 3);

                // Delete a column that contains the "B2"cell.
                worksheet.DeleteCells(worksheet.Cells["B2"], DeleteMode.EntireColumn);
                #endregion #DeleteColumns
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void CopyRowsColumns(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets[0];

                // Modify the 2nd row. 
                worksheet.Cells["A2"].Value = "Row 2";
                worksheet.Rows["2"].Height = 50;
                worksheet.Rows["2"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows["2"].FillColor = Color.LightCyan;

                // Modify the "B" column.
                worksheet.Cells["B1"].Value = "ColumnB";
                worksheet.Columns["B"].Borders.SetOutsideBorders(Color.CadetBlue, BorderLineStyle.Thick);

                #region #CopyRowsColumns
                // Copy all data from the 2nd row to the 5th row.
                worksheet.Rows["5"].CopyFrom(worksheet.Rows["2"]);

                // Copy only borders from the "B" column to the "E" column.
                worksheet.Columns["E"].CopyFrom(worksheet.Columns["B"], PasteSpecial.Borders);
                #endregion #CopyRowsColumns
            }
            finally {
                workbook.EndUpdate();
            }
        }

        static void ShowHideRowsColumns(IWorkbook workbook) {
            #region #ShowHideRowsColumns
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the 8th row of the worksheet.
            worksheet.Rows[7].Visible = false;

            // Hide the 4th column of the worksheet.
            worksheet.Columns[3].Visible = false;
            #endregion #ShowHideRowsColumns
        }

        static void SpecifyRowsHeightColumnsWidth(IWorkbook workbook) {
            workbook.BeginUpdate();
            try {

            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["B1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Cells["B1"].Value = "30 characters";
            worksheet.Cells["C1"].Value = "15 mm";
            worksheet.Cells["E1"].Value = "100 pt";
            worksheet.Cells["F1"].Value = "70 pt";
            worksheet.Cells["G1"].Value = "70 pt";
            worksheet.Cells["H1"].Value = "70 pt";
            worksheet.Cells["J1"].Value = "30 characters";
            worksheet.Cells["K1"].Value = "15 mm";

            worksheet.Cells["A3"].Value = "50 pt";
            worksheet.Cells["A5"].Value = "2\"";
            worksheet.Cells["A7"].Value = "50 pt";
            Formatting rowHeightValues = worksheet.Range["A3:A7"].BeginUpdateFormatting();
            rowHeightValues.Alignment.RotationAngle = 90;
            rowHeightValues.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            rowHeightValues.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Range["A3:A7"].EndUpdateFormatting(rowHeightValues);

            #region #RowHeight
            // Set the height of the 3rd row to 50 points
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.Rows[2].Height = 50;

            // Set the height of the row that contains the "C5" cell to 2 inches.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;
            worksheet.Cells["C5"].RowHeight = 2;

            // Set the height of the 7th row to the height of the 3rd row.
            worksheet.Rows["7"].Height = worksheet.Rows["3"].Height;

            // Set the default row height to 30 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.DefaultRowHeight = 30;
            #endregion #RowHeight

            #region #ColumnWidth
            // Set the "B" column width to 30 characters of the default font that is specified by the Normal style.
            worksheet.Columns["B"].WidthInCharacters = 30;

            // Set the "C" column width to 15 millimeters.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            worksheet.Columns["C"].Width = 15;

            // Set the width of the column that contains the "E15" cell to 100 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.Cells["E15"].ColumnWidth = 100;

            // Set the width of all columns that contain the "F4:H7" cell range (the "F", "G" and "H" columns) to 70 points.
            worksheet.Range["F4:H7"].ColumnWidth = 70;

            // Set the "J" column width to the "B" column width value.
            worksheet.Columns["J"].Width = worksheet.Columns["B"].Width;

            // Copy the "C" column width value and assign it to the "K" column width.
            worksheet.Columns["K"].CopyFrom(worksheet.Columns["C"], PasteSpecial.ColumnWidths);

            // Set the default column width to 40 pixels.
            worksheet.DefaultColumnWidthInPixels = 40;
            #endregion #ColumnWidth
            }
            finally {
                workbook.EndUpdate();
            }
        }

    }
}
