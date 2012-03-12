Imports System.Reflection

Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.VBIDEApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' initialize api
        LateBindingApi.Core.Factory.Initialize()

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application
        wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        ' add a new document
        Dim newDocument As Word.Document
        newDocument = wordApplication.Documents.Add()

        ' add a table
        Dim table As Word.Table
        table = newDocument.Tables.Add(wordApplication.Selection.Range, 3, 2)

        'insert some text into the cells
        table.Cell(1, 1).Select()
        wordApplication.Selection.TypeText("This")

        table.Cell(1, 2).Select()
        wordApplication.Selection.TypeText("table")

        table.Cell(2, 1).Select()
        wordApplication.Selection.TypeText("was")

        table.Cell(2, 2).Select()
        wordApplication.Selection.TypeText("created")

        table.Cell(3, 1).Select()
        wordApplication.Selection.TypeText("by")

        table.Cell(3, 2).Select()
        wordApplication.Selection.TypeText("NetOffice")

        'save the document
        Dim fileExtension As String = GetDefaultExtension(wordApplication)
        Dim documentFile As String = String.Format("{0}\Example02{1}", Application.StartupPath, fileExtension)

        ' newer word versions try to save the document in .odt(open document format) by default
        ' we dont want this, we want .doc or .docx !!!
        Dim version As Double = Convert.ToDouble(wordApplication.Version)
        If (version >= 120.0) Then
            newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault)
        Else
            newDocument.SaveAs(documentFile)
        End If

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        Dim fDialog As New FinishDialog("Document saved.", documentFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".doc" or ".docx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Word.Application) As String

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".docx"
        Else
            Return ".doc"
        End If

    End Function

#End Region

End Class
