Imports System.Collections.Generic

Namespace SpreadsheetControl_WPF_API

    Public Partial Class Groups
        Inherits List(Of Group)

        Public Shared Function InitData() As Groups
            Dim examples As Groups = New Groups()
'#Region "GroupNodes"
            examples.Add(New Group("Worksheet"))
            examples.Add(New Group("Rows and Columns"))
            examples.Add(New Group("Cells"))
            examples.Add(New Group("Formulas"))
            examples.Add(New Group("Formatting"))
            examples.Add(New Group("Import/Export"))
            examples.Add(New Group("Printing"))
'#End Region
'#Region "ExampleNodes"
            ' Add nodes to the "Worksheet" group of examples.
            examples(0).Items.Add(New SpreadsheetExample("Active Worksheet", AssignActiveWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("New Worksheet", AddWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("Delete a Worksheet", RemoveWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("Rename a Worksheet", RenameWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("Copy a Worksheet within a Workbook", CopyWorksheetWithinWorkbookAction))
            examples(0).Items.Add(New SpreadsheetExample("Move a Worksheet", MoveWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("Show/Hide a Worksheet", ShowHideWorksheetAction))
            examples(0).Items.Add(New SpreadsheetExample("Show/Hide Gridlines", ShowHideGridlinesAction))
            examples(0).Items.Add(New SpreadsheetExample("Show/Hide Row and Column Headings", ShowHideHeadingsAction))
            examples(0).Items.Add(New SpreadsheetExample("Zoom a Worksheet", ZoomWorksheetAction))
            ' Add nodes to the "Rows and Columns" group of examples.
            examples(1).Items.Add(New SpreadsheetExample("New Row/Column", InsertRowsColumnsAction))
            examples(1).Items.Add(New SpreadsheetExample("Delete a Row/Column", DeleteRowsColumnsAction))
            examples(1).Items.Add(New SpreadsheetExample("Copy a Row/Column", CopyRowsColumnsAction))
            examples(1).Items.Add(New SpreadsheetExample("Show or Hide a Row/Column", ShowHideRowsColumnsAction))
            examples(1).Items.Add(New SpreadsheetExample("Row Height and Column Width", SpecifyRowsHeightColumnsWidthAction))
            ' Add nodes to the "Cells" group of examples.
            examples(2).Items.Add(New SpreadsheetExample("Cell Value", ChangeCellValueAction))
            examples(2).Items.Add(New SpreadsheetExample("Cell Value To/From Object", CellValueToFromObjectAction))
            examples(2).Items.Add(New SpreadsheetExample("Cell Value From Object via Custom Converter", CellValueFromObjectViaCustomConverterAction))
            examples(2).Items.Add(New SpreadsheetExample("Add Hyperlinks to Cells", AddHyperlinkAction))
            examples(2).Items.Add(New SpreadsheetExample("Create, Edit and Copy Comments", AddCommentAction))
            examples(2).Items.Add(New SpreadsheetExample("Copy Data Only, Style Only, or Data with Style", CopyCellDataAndStyleAction))
            examples(2).Items.Add(New SpreadsheetExample("Merge/Split Cells", MergeAndSplitCellsAction))
            examples(2).Items.Add(New SpreadsheetExample("Clear Cells", ClearCellsAction))
            ' Add nodes to the "Formulas" group of examples. 
            examples(3).Items.Add(New SpreadsheetExample("Constants and Calculation Operators in Formulas", UseConstantsAndCalculationOperatorsInFormulasAction))
            examples(3).Items.Add(New SpreadsheetExample("R1C1 References in Formulas", R1C1ReferencesInFormulassAction))
            examples(3).Items.Add(New SpreadsheetExample("Names in Formulas", UseNamesInFormulasAction))
            examples(3).Items.Add(New SpreadsheetExample("Create Named Formulas", CreateNamedFormulasAction))
            examples(3).Items.Add(New SpreadsheetExample("Functions in Formulas", UseFunctionsInFormulasAction))
            examples(3).Items.Add(New SpreadsheetExample("Shared and Array Formulas", CreateSharedAndArrayFormulasAction))
            ' Add nodes to the "Formatting" group of examples.
            examples(4).Items.Add(New SpreadsheetExample("Create, Modify and Apply a Style", FormattingActions.CreateModifyApplyStyleAction))
            examples(4).Items.Add(New SpreadsheetExample("Cell and Cell Range Formatting", FormattingActions.FormatCellAction))
            examples(4).Items.Add(New SpreadsheetExample("Date Formats", FormattingActions.SetDateFormatsAction))
            examples(4).Items.Add(New SpreadsheetExample("Number Formats", FormattingActions.SetNumberFormatsAction))
            examples(4).Items.Add(New SpreadsheetExample("Custom Number Format", FormattingActions.CustomNumberFormatAction))
            examples(4).Items.Add(New SpreadsheetExample("Cell Colors and Background", FormattingActions.ChangeCellColorsAction))
            examples(4).Items.Add(New SpreadsheetExample("Cell Gradient Fill", FormattingActions.ChangeCellGradientFillAction))
            examples(4).Items.Add(New SpreadsheetExample("Font Settings", FormattingActions.SpecifyCellFontAction))
            examples(4).Items.Add(New SpreadsheetExample("Cell Alignment", FormattingActions.AlignCellContentsAction))
            examples(4).Items.Add(New SpreadsheetExample("Cell Borders", FormattingActions.AddCellBordersAction))
            ' Add nodes to the "Import/Export" group of examples.
            examples(5).Items.Add(New SpreadsheetExample("Export to Pdf", ExportToPdfAction))
            ' Add nodes to the "Printing" group of examples.
            examples(6).Items.Add(New SpreadsheetExample("Print a Workbook", PrintAction))
            Return examples
'#End Region
        End Function
    End Class
End Namespace
