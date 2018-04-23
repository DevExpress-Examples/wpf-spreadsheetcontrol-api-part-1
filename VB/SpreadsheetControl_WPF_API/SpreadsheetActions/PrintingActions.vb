Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
#Region "#printingUsings"
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraPrinting
Imports DevExpress.Xpf.Printing
Imports System.Windows.Controls
#End Region ' #printingUsings

Namespace SpreadsheetControl_WPF_API
	Public NotInheritable Class PrintingActions

        Public Shared PrintAction As Action(Of IWorkbook) = AddressOf Print

		Private Sub New()
		End Sub
		Private Shared Sub Print(ByVal workbook As IWorkbook)

			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Generate worksheet content - the simple multiplication table.
			Dim columnHeadings As Range = worksheet.Range.FromLTRB(1,0,40,0)
			columnHeadings.Formula = "=COLUMN() - 1"
			Dim rowHeadings As Range = worksheet.Range.FromLTRB(0,1,0,40)
			rowHeadings.Formula = "=ROW() - 1"
			Dim tableRange As Range = worksheet.Range.FromLTRB(1,1,40,40)
			tableRange.Formula = "=(ROW()-1)*(COLUMN()-1)"

			' Format headers of the multiplication table.
			Dim rangeFormatting As Formatting = columnHeadings.BeginUpdateFormatting()
			rangeFormatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
			rangeFormatting.Borders.BottomBorder.Color = Color.Black
			columnHeadings.EndUpdateFormatting(rangeFormatting)

			rangeFormatting = rowHeadings.BeginUpdateFormatting()
			rangeFormatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin
			rangeFormatting.Borders.RightBorder.Color = Color.Black
			rowHeadings.EndUpdateFormatting(rangeFormatting)

			rangeFormatting = tableRange.BeginUpdateFormatting()
			rangeFormatting.Fill.BackgroundColor = Color.LightBlue
			tableRange.EndUpdateFormatting(rangeFormatting)

'			#Region "#WorksheetPrintOptions"
			worksheet.ActiveView.Orientation = PageOrientation.Landscape
			'  Display row and column headings.
			worksheet.ActiveView.ShowHeadings = True
			worksheet.ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4
			' Access an object that contains print options.
			Dim printOptions As WorksheetPrintOptions = worksheet.PrintOptions
			'  Print in black and white.
			printOptions.BlackAndWhite = True
			'  Do not print gridlines.
			printOptions.PrintGridlines = False
			'  Scale the print area to fit to two pages wide.
			printOptions.FitToPage = True
			printOptions.FitToWidth = 2
			'  Print a dash instead of a cell error message.
			printOptions.ErrorsPrintMode = ErrorsPrintMode.Dash
'			#End Region ' #WorksheetPrintOptions

'			#Region "#PrintWorkbook"
			' Invoke the Print Preview dialog for the workbook.
			Using link As New LegacyPrintableComponentLink(workbook)
				link.CreateDocument()
				link.ShowPrintPreviewDialog(App.Current.MainWindow)
			End Using
'			#End Region ' #PrintWorkbook
		End Sub
	End Class
End Namespace
