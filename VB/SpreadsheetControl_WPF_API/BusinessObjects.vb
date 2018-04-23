Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation
Imports System.Windows.Shapes
Imports DevExpress.Xpf.NavBar
Imports DevExpress.Spreadsheet


Namespace SpreadsheetControl_WPF_API
    Public Class Group
        Public Property Header() As String
        Public Property Items() As List(Of SpreadsheetExample)

        Public Sub New(ByVal name As String)
            Header = name
            Items = New List(Of SpreadsheetExample)()
        End Sub
    End Class

    Public Class SpreadsheetExample
        Public Property Header() As String
        Public Sub New(ByVal name As String, ByVal action As Action(Of IWorkbook))
            Header = name
            Me.Action = action
        End Sub
        Private privateAction As Action(Of IWorkbook)
        Public Property Action() As Action(Of IWorkbook)
            Get
                Return privateAction
            End Get
            Private Set(ByVal value As Action(Of IWorkbook))
                privateAction = value
            End Set
        End Property
    End Class
End Namespace
