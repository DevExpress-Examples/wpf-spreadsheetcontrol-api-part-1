Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.NavBar
Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace SpreadsheetControl_WPF_API

    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Public Partial Class MainWindow
        Inherits Window

        Private workbook As IWorkbook

        Public Sub New()
            Me.InitializeComponent()
            ' Access a workbook.
            workbook = Me.spreadsheetControl1.Document
            DataContext = Groups.InitData()
        End Sub

        Private Sub NavigationPaneView_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            Dim group As NavBarGroup = CType(sender, NavBarViewBase).GetNavBarGroup(e)
            Dim item As NavBarItem = CType(sender, NavBarViewBase).GetNavBarItem(e)
            If item Is Nothing Then Return
            Dim example As SpreadsheetExample = TryCast(item.Content, SpreadsheetExample)
            If example Is Nothing Then Return
            LoadDocumentFromFile()
            Dim action As Action(Of IWorkbook) = example.Action
            action(workbook)
            SaveDocumentToFile()
        End Sub

        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
'#Region "#LoadDocumentFromFile"
            ' Load a workbook from a file.
            workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #LoadDocumentFromFile
        End Sub

        Private Sub SaveDocumentToFile()
'#Region "#SaveDocumentToFile"
            ' Save the modified document to a file.
            workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'#End Region  ' #SaveDocumentToFile
        End Sub
    End Class
End Namespace
