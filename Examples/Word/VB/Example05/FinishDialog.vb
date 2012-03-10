Imports System.Diagnostics

Public NotInheritable Class FinishDialog

    Dim _message As String
    Dim _documentPath As String

    Private Sub FinishDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub buttonClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonClose.Click
        Me.Close()
    End Sub

    Private Sub buttonOpenWorkbook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonOpenDocument.Click
        Process.Start(_documentPath)
        Me.Close()
    End Sub

    Public Sub New(ByVal message As String, ByVal workbookPath As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _message = message
        _documentPath = workbookPath

        labelMessage.Text = _message
        labelDocumentPath.Text = _documentPath

    End Sub

End Class
