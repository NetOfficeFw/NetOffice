Imports System.IO
Imports System.Reflection

Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.VBIDEApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' create simple a csv-file as datasource
        Dim fileName As String = String.Format("{0}\DataSource.csv", Application.StartupPath)

        ' if file exists then delete
        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        File.AppendAllText(fileName, String.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine))
        File.AppendAllText(fileName, String.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com/", Environment.NewLine))

        ' initialize api
        LateBindingApi.Core.Factory.Initialize()

        ' start word and turn off msg boxes
        Dim wordApplication As New Word.Application
        wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        ' add a new document
        Dim newDocument As Word.Document
        newDocument = wordApplication.Documents.Add()

        ' define the document as mailmerge
        newDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters

        ' open the datasource
        newDocument.MailMerge.OpenDataSource(fileName)

        ' insert some text and the mailmergefields defined in the datasource
        wordApplication.Selection.TypeText("This test is brought to you by ")
        newDocument.MailMerge.Fields.Add(wordApplication.Selection.Range, "ProjectName")

        wordApplication.Selection.TypeText(" for more information and examples visit ")
        newDocument.MailMerge.Fields.Add(wordApplication.Selection.Range, "ProjectLink ")

        wordApplication.Selection.TypeText(" or click ")

        newDocument.Hyperlinks.Add(wordApplication.Selection.Range, newDocument.MailMerge.DataSource.DataFields(2).Value, _
                                   Missing.Value, "come on dude click me, i know you want it...", "here", Missing.Value)

        'show the contents of the fields
        newDocument.MailMerge.ViewMailMergeFieldCodes = WdConstants.wdToggle

        'do not show the fieldcodes
        wordApplication.ActiveWindow.View.ShowFieldCodes = False

        'save the document
        Dim fileExtension As String = GetDefaultExtension(wordApplication)
        Dim documentFile As String = String.Format("{0}\Example04{1}", Application.StartupPath, fileExtension)
        newDocument.SaveAs(documentFile)

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
