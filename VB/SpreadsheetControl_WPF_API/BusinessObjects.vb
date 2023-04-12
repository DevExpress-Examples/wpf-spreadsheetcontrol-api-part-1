Imports System
Imports System.Collections.Generic
Imports System.Windows.Controls
Imports DevExpress.Spreadsheet

Namespace SpreadsheetControl_WPF_API

    Public Class Group

        Public Property Header As String

        Public Property Items As List(Of SpreadsheetExample)

        Public Sub New(ByVal name As String)
            Header = name
            Items = New List(Of SpreadsheetExample)()
        End Sub
    End Class

    Public Class SpreadsheetExample

        Private _Action As Action(Of DevExpress.Spreadsheet.IWorkbook)

        Public Property Header As String

        Public Sub New(ByVal name As String, ByVal action As Action(Of IWorkbook))
            Header = name
            Me.Action = action
        End Sub

        Public Property Action As Action(Of IWorkbook)
            Get
                Return _Action
            End Get

            Private Set(ByVal value As Action(Of IWorkbook))
                _Action = value
            End Set
        End Property
    End Class
End Namespace
