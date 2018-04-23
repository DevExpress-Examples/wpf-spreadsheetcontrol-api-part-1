Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetControl_WPF_API
    Public NotInheritable Class WorksheetActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared AssignActiveWorksheetAction As Action(Of IWorkbook) = AddressOf AssignActiveWorksheet
        Public Shared AddWorksheetAction As Action(Of IWorkbook) = AddressOf AddWorksheet
        Public Shared RemoveWorksheetAction As Action(Of IWorkbook) = AddressOf RemoveWorksheet
        Public Shared RenameWorksheetAction As Action(Of IWorkbook) = AddressOf RenameWorksheet
        Public Shared CopyWorksheetWithinWorkbookAction As Action(Of IWorkbook) = AddressOf CopyWorksheetWithinWorkbook
        Public Shared MoveWorksheetAction As Action(Of IWorkbook) = AddressOf MoveWorksheet
        Public Shared ShowHideWorksheetAction As Action(Of IWorkbook) = AddressOf ShowHideWorksheet
        Public Shared ShowHideGridlinesAction As Action(Of IWorkbook) = AddressOf ShowHideGridlines
        Public Shared ShowHideHeadingsAction As Action(Of IWorkbook) = AddressOf ShowHideHeadings
        Public Shared ZoomWorksheetAction As Action(Of IWorkbook) = AddressOf ZoomWorksheet
        #End Region

        Private Shared Sub AssignActiveWorksheet(ByVal workbook As IWorkbook)
'            #Region "#ActiveWorksheet"
            ' Set the second worksheet ("Sheet2") as the active worksheet.
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets("Sheet2")
'            #End Region ' #ActiveWorksheet
        End Sub

        Private Shared Sub AddWorksheet(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
'                #Region "#AddWorksheet"
                ' Add a new worksheet to the workbook. 
                ' The worksheet will be inserted into the end of the existing collection of worksheets.
                ' Worksheet name is "SheetN", where N is a number following the largest number used in existing worksheet names of the same type.
                workbook.Worksheets.Add()

                ' Add a new worksheet under the specified name.
                workbook.Worksheets.Add().Name = "TestSheet1"

                workbook.Worksheets.Add("TestSheet2")

                ' Add a new worksheet at the specified position in the collection of worksheets.
                workbook.Worksheets.Insert(1, "TestSheet3")

                workbook.Worksheets.Insert(3)

'                #End Region ' #AddWorksheet
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub RemoveWorksheet(ByVal workbook As IWorkbook)
'            #Region "#DeleteWorksheet"
            ' Delete the "Sheet2" worksheet from the workbook.
            workbook.Worksheets.Remove(workbook.Worksheets("Sheet2"))

            ' Delete the first worksheet from the workbook.
            workbook.Worksheets.RemoveAt(0)
'            #End Region ' #DeleteWorksheet
        End Sub

        Private Shared Sub RenameWorksheet(ByVal workbook As IWorkbook)
'            #Region "#RenameWorksheet"
            ' Change the name of the second worksheet.
            workbook.Worksheets(1).Name = "Renamed Sheet"
'            #End Region ' #RenameWorksheet
        End Sub

        Private Shared Sub CopyWorksheetWithinWorkbook(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                workbook.Worksheets("Sheet1").Cells.FillColor = Color.LightSteelBlue
                workbook.Worksheets("Sheet1").Cells("A1").ColumnWidthInCharacters = 50
                workbook.Worksheets("Sheet1").Cells("A1").Value = "Sheet1's Content"

'                #Region "#CopyWorksheet"
                ' Add a new worksheet to a workbook.
                workbook.Worksheets.Add("Sheet1_Copy")

                ' Copy all information (content and formatting) to the newly created worksheet 
                ' from the "Sheet1" worksheet.
                workbook.Worksheets("Sheet1_Copy").CopyFrom(workbook.Worksheets("Sheet1"))
'                #End Region ' #CopyWorksheet
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Shared Sub MoveWorksheet(ByVal workbook As IWorkbook)
'            #Region "#MoveWorksheet"
            ' Move the first worksheet to the position of the last worksheet within the workbook.
            workbook.Worksheets(0).Move(workbook.Worksheets.Count - 1)
'            #End Region ' #MoveWorksheet
        End Sub

        Private Shared Sub ShowHideWorksheet(ByVal workbook As IWorkbook)
'            #Region "#ShowHideWorksheet"
            ' Hide the "Sheet2" worksheet and prevent end-users from unhiding it via the user interface.
            ' To make this worksheet visible again, use the Worksheet.Visible property.
            workbook.Worksheets("Sheet2").VisibilityType = WorksheetVisibilityType.VeryHidden

            ' Hide the "Sheet3" worksheet. 
            ' In this state, the worksheet can be unhidden via the user interface.
            workbook.Worksheets("Sheet3").Visible = False
'            #End Region ' #ShowHideWorksheet
        End Sub

        Private Shared Sub ShowHideGridlines(ByVal workbook As IWorkbook)
'            #Region "#ShowHideGridlines"
            ' Hide gridlines in the first worksheet.
            workbook.Worksheets(0).ActiveView.ShowGridlines = False
'            #End Region ' #ShowHideGridlines
        End Sub

        Private Shared Sub ShowHideHeadings(ByVal workbook As IWorkbook)
'            #Region "#ShowHideHeadings"
            ' Hide row and column headings in the first worksheet.
            workbook.Worksheets(0).ActiveView.ShowHeadings = False
'            #End Region ' #ShowHideHeadings
        End Sub

        Private Shared Sub ZoomWorksheet(ByVal workbook As IWorkbook)
'            #Region "#WorksheetZoom"
            ' Zoom in to the worksheet view. 
            workbook.Worksheets(0).ActiveView.Zoom = 150
'            #End Region ' #WorksheetZoom
        End Sub
    End Class
End Namespace
