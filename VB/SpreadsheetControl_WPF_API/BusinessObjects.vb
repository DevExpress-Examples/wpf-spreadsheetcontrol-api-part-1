Imports Microsoft.VisualBasic
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
		Private privateHeader As String
		Public Property Header() As String
			Get
				Return privateHeader
			End Get
			Set(ByVal value As String)
				privateHeader = value
			End Set
		End Property
		Private privateItems As List(Of SpreadsheetExample)
		Public Property Items() As List(Of SpreadsheetExample)
			Get
				Return privateItems
			End Get
			Set(ByVal value As List(Of SpreadsheetExample))
				privateItems = value
			End Set
		End Property

		Public Sub New(ByVal name As String)
			Header = name
			Items = New List(Of SpreadsheetExample)()
		End Sub
	End Class

	Public Class SpreadsheetExample
        Private privateHeader As String
        Private privateAction As Action(Of IWorkbook)

		Public Property Header() As String
			Get
				Return privateHeader
			End Get
			Set(ByVal value As String)
				privateHeader = value
			End Set
		End Property
        Public Sub New(ByVal name As String, ByVal action As Action(Of IWorkbook))
            Header = name
            privateAction = action
        End Sub

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
