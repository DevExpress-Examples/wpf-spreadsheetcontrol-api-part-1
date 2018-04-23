using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetControl_WPF_API
{
    public static class WorksheetActions
    {
        #region Actions
        public static Action<IWorkbook> AssignActiveWorksheetAction = AssignActiveWorksheet;
        public static Action<IWorkbook> AddWorksheetAction = AddWorksheet;
        public static Action<IWorkbook> RemoveWorksheetAction = RemoveWorksheet;
        public static Action<IWorkbook> RenameWorksheetAction = RenameWorksheet;
        public static Action<IWorkbook> CopyWorksheetWithinWorkbookAction = CopyWorksheetWithinWorkbook;
        public static Action<IWorkbook> MoveWorksheetAction = MoveWorksheet;
        public static Action<IWorkbook> ShowHideWorksheetAction = ShowHideWorksheet;
        public static Action<IWorkbook> ShowHideGridlinesAction = ShowHideGridlines;
        public static Action<IWorkbook> ShowHideHeadingsAction = ShowHideHeadings;
        public static Action<IWorkbook> ZoomWorksheetAction = ZoomWorksheet;
        #endregion

        static void AssignActiveWorksheet(IWorkbook workbook)
        {
            #region #ActiveWorksheet
            // Set the second worksheet ("Sheet2") as the active worksheet.
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["Sheet2"];
            #endregion #ActiveWorksheet
        }

        static void AddWorksheet(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                #region #AddWorksheet
                // Add a new worksheet to the workbook. 
                // The worksheet will be inserted into the end of the existing collection of worksheets.
                // Worksheet name is "SheetN", where N is a number following the largest number used in existing worksheet names of the same type.
                workbook.Worksheets.Add();

                // Add a new worksheet under the specified name.
                workbook.Worksheets.Add().Name = "TestSheet1";

                workbook.Worksheets.Add("TestSheet2");

                // Add a new worksheet at the specified position in the collection of worksheets.
                workbook.Worksheets.Insert(1, "TestSheet3");

                workbook.Worksheets.Insert(3);

                #endregion #AddWorksheet
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void RemoveWorksheet(IWorkbook workbook)
        {
            #region #DeleteWorksheet
            // Delete the "Sheet2" worksheet from the workbook.
            workbook.Worksheets.Remove(workbook.Worksheets["Sheet2"]);

            // Delete the first worksheet from the workbook.
            workbook.Worksheets.RemoveAt(0);
            #endregion #DeleteWorksheet
        }

        static void RenameWorksheet(IWorkbook workbook)
        {
            #region #RenameWorksheet
            // Change the name of the second worksheet.
            workbook.Worksheets[1].Name = "Renamed Sheet";
            #endregion #RenameWorksheet
        }

        static void CopyWorksheetWithinWorkbook(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                workbook.Worksheets["Sheet1"].Cells.FillColor = Color.LightSteelBlue;
                workbook.Worksheets["Sheet1"].Cells["A1"].ColumnWidthInCharacters = 50;
                workbook.Worksheets["Sheet1"].Cells["A1"].Value = "Sheet1's Content";

                #region #CopyWorksheet
                // Add a new worksheet to a workbook.
                workbook.Worksheets.Add("Sheet1_Copy");

                // Copy all information (content and formatting) to the newly created worksheet 
                // from the "Sheet1" worksheet.
                workbook.Worksheets["Sheet1_Copy"].CopyFrom(workbook.Worksheets["Sheet1"]);
                #endregion #CopyWorksheet
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void MoveWorksheet(IWorkbook workbook)
        {
            #region #MoveWorksheet
            // Move the first worksheet to the position of the last worksheet within the workbook.
            workbook.Worksheets[0].Move(workbook.Worksheets.Count - 1);
            #endregion #MoveWorksheet
        }

        static void ShowHideWorksheet(IWorkbook workbook)
        {
            #region #ShowHideWorksheet
            // Hide the "Sheet2" worksheet and prevent end-users from unhiding it via the user interface.
            // To make this worksheet visible again, use the Worksheet.Visible property.
            workbook.Worksheets["Sheet2"].VisibilityType = WorksheetVisibilityType.VeryHidden;

            // Hide the "Sheet3" worksheet. 
            // In this state, the worksheet can be unhidden via the user interface.
            workbook.Worksheets["Sheet3"].Visible = false;
            #endregion #ShowHideWorksheet
        }

        static void ShowHideGridlines(IWorkbook workbook)
        {
            #region #ShowHideGridlines
            // Hide gridlines in the first worksheet.
            workbook.Worksheets[0].ActiveView.ShowGridlines = false;
            #endregion #ShowHideGridlines
        }

        static void ShowHideHeadings(IWorkbook workbook)
        {
            #region #ShowHideHeadings
            // Hide row and column headings in the first worksheet.
            workbook.Worksheets[0].ActiveView.ShowHeadings = false;
            #endregion #ShowHideHeadings
        }

        static void ZoomWorksheet(IWorkbook workbook)
        {
            #region #WorksheetZoom
            // Zoom in to the worksheet view. 
            workbook.Worksheets[0].ActiveView.Zoom = 150;
            #endregion #WorksheetZoom
        }
    }
}
