using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetControl_WPF_API
{
    public partial class Groups : List<Group>
    {
        public static Groups InitData()
        {
            Groups examples = new Groups();

            #region GroupNodes
            examples.Add(new Group("Worksheet"));
            examples.Add(new Group("Rows and Columns"));
            examples.Add(new Group("Cells"));
            examples.Add(new Group("Formulas"));
            examples.Add(new Group("Formatting"));
            examples.Add(new Group("Import/Export"));
            examples.Add(new Group("Printing"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Worksheet" group of examples.
            examples[0].Items.Add(new SpreadsheetExample("Active Worksheet", WorksheetActions.AssignActiveWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("New Worksheet", WorksheetActions.AddWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("Delete a Worksheet", WorksheetActions.RemoveWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("Rename a Worksheet", WorksheetActions.RenameWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("Copy a Worksheet within a Workbook", WorksheetActions.CopyWorksheetWithinWorkbookAction));
            examples[0].Items.Add(new SpreadsheetExample("Move a Worksheet", WorksheetActions.MoveWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("Show/Hide a Worksheet", WorksheetActions.ShowHideWorksheetAction));
            examples[0].Items.Add(new SpreadsheetExample("Show/Hide Gridlines", WorksheetActions.ShowHideGridlinesAction));
            examples[0].Items.Add(new SpreadsheetExample("Show/Hide Row and Column Headings", WorksheetActions.ShowHideHeadingsAction));
            examples[0].Items.Add(new SpreadsheetExample("Zoom a Worksheet", WorksheetActions.ZoomWorksheetAction));

            // Add nodes to the "Rows and Columns" group of examples.
            examples[1].Items.Add(new SpreadsheetExample("New Row/Column", RowAndColumnActions.InsertRowsColumnsAction));
            examples[1].Items.Add(new SpreadsheetExample("Delete a Row/Column", RowAndColumnActions.DeleteRowsColumnsAction));
            examples[1].Items.Add(new SpreadsheetExample("Copy a Row/Column", RowAndColumnActions.CopyRowsColumnsAction));
            examples[1].Items.Add(new SpreadsheetExample("Show or Hide a Row/Column", RowAndColumnActions.ShowHideRowsColumnsAction));
            examples[1].Items.Add(new SpreadsheetExample("Row Height and Column Width", RowAndColumnActions.SpecifyRowsHeightColumnsWidthAction));

            // Add nodes to the "Cells" group of examples.
            examples[2].Items.Add(new SpreadsheetExample("Cell Value", CellActions.ChangeCellValueAction));
            examples[2].Items.Add(new SpreadsheetExample("Add Hyperlinks to Cells", CellActions.AddHyperlinkAction));
            examples[2].Items.Add(new SpreadsheetExample("Copy Data Only, Style Only, or Data with Style", CellActions.CopyCellDataAndStyleAction));
            examples[2].Items.Add(new SpreadsheetExample("Merge/Split Cells", CellActions.MergeAndSplitCellsAction));
            examples[2].Items.Add(new SpreadsheetExample("Clear Cells", CellActions.ClearCellsAction));

            // Add nodes to the "Formulas" group of examples. 
            examples[3].Items.Add(new SpreadsheetExample("Constants and Calculation Operators in Formulas", FormulaActions.UseConstantsAndCalculationOperatorsInFormulasAction));
            examples[3].Items.Add(new SpreadsheetExample("R1C1 References in Formulas", FormulaActions.R1C1ReferencesInFormulassAction));
            examples[3].Items.Add(new SpreadsheetExample("Names in Formulas", FormulaActions.UseNamesInFormulasAction));
            examples[3].Items.Add(new SpreadsheetExample("Create Named Formulas", FormulaActions.CreateNamedFormulasAction));
            examples[3].Items.Add(new SpreadsheetExample("Functions in Formulas", FormulaActions.UseFunctionsInFormulasAction));
            examples[3].Items.Add(new SpreadsheetExample("Shared and Array Formulas", FormulaActions.CreateSharedAndArrayFormulasAction));

            // Add nodes to the "Formatting" group of examples.
            examples[4].Items.Add(new SpreadsheetExample("Create, Modify and Apply a Style", FormattingActions.CreateModifyApplyStyleAction));
            examples[4].Items.Add(new SpreadsheetExample("Cell and Cell Range Formatting", FormattingActions.FormatCellAction));
            examples[4].Items.Add(new SpreadsheetExample("Date Formats", FormattingActions.SetDateFormatsAction));
            examples[4].Items.Add(new SpreadsheetExample("Number Formats", FormattingActions.SetNumberFormatsAction));
            examples[4].Items.Add(new SpreadsheetExample("Custom Number Format", FormattingActions.CustomNumberFormatAction));
            examples[4].Items.Add(new SpreadsheetExample("Cell Colors and Background", FormattingActions.ChangeCellColorsAction));
            examples[4].Items.Add(new SpreadsheetExample("Cell Gradient Fill", FormattingActions.ChangeCellGradientFillAction));
            examples[4].Items.Add(new SpreadsheetExample("Font Settings", FormattingActions.SpecifyCellFontAction));
            examples[4].Items.Add(new SpreadsheetExample("Cell Alignment", FormattingActions.AlignCellContentsAction));
            examples[4].Items.Add(new SpreadsheetExample("Cell Borders", FormattingActions.AddCellBordersAction));

            // Add nodes to the "Import/Export" group of examples.
            examples[5].Items.Add(new SpreadsheetExample("Export to Pdf", ImportExportActions.ExportToPdfAction));

            // Add nodes to the "Printing" group of examples.
            examples[6].Items.Add(new SpreadsheetExample("Print a Workbook", PrintingActions.PrintAction));

            return examples;
            #endregion
        }
    }
}
