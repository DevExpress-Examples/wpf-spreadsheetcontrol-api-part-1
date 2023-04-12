Imports System
Imports System.IO
Imports System.Diagnostics
Imports DevExpress.Spreadsheet

Namespace SpreadsheetControl_WPF_API

    Public Module ImportExportActions

        Public ExportToPdfAction As Action(Of IWorkbook) = AddressOf ExportToPdf

        Private Sub ExportToPdf(ByVal workbook As IWorkbook)
            workbook.Worksheets(0).Cells("D8").Value = "This document is exported to the PDF format."
'#Region "#ExportToPdf"
            Using pdfFileStream As FileStream = New FileStream("Documents\Document_PDF.pdf", FileMode.Create)
                workbook.ExportToPdf(pdfFileStream)
            End Using

'#End Region  ' #ExportToPdf
            Call Process.Start("Documents\Document_PDF.pdf")
        End Sub
    End Module
End Namespace
