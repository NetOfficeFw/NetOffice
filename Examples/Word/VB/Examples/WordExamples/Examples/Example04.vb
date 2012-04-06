Imports ExampleBase
Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums

Public Class Example04
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' create simple a csv-file as datasource
        Dim fileName As String = String.Format("{0}\DataSource.csv", _hostApplication.RootDirectory)

        ' if file exists then delete
        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        File.AppendAllText(fileName, String.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine))
        File.AppendAllText(fileName, String.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com/", Environment.NewLine))

        
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
                                   Type.Missing, "click me if know you want.", "here", Type.Missing)

        'show the contents of the fields
        newDocument.MailMerge.ViewMailMergeFieldCodes = WdConstants.wdToggle

        ' do not show the fieldcodes
        wordApplication.ActiveWindow.View.ShowFieldCodes = False

        ' we save the document as .doc for compatibility with all word versions
        Dim documentFile As String = String.Format("{0}\Example03{1}", _hostApplication.RootDirectory, ".doc")
        newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault)

        ' close word and dispose reference
        wordApplication.Quit()
        wordApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example04", "Beispiel04")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Using data source", "Verwendung von DataSource")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

End Class
