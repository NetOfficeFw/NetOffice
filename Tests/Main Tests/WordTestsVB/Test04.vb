Imports NetOffice
Imports System.Reflection
Imports System.IO
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test04
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using List templates."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test04"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Word"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Word.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.WordApi.Application()
            application.DisplayAlerts = WdAlertLevel.wdAlertsNone

            ' add a new document
            Dim newDocument As Word.Document
            newDocument = application.Documents.Add()

            ' create simple a csv-file as datasource
            Dim fileName As String = String.Format("{0}\\DataSource.csv", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))

            'if file exists then delete
            If (File.Exists(fileName)) Then
                File.Delete(fileName)
            End If

            File.AppendAllText(fileName, String.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine))
            File.AppendAllText(fileName, String.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com/", Environment.NewLine))

            ' define the document as mailmerge
            newDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters

            ' open the datasource
            newDocument.MailMerge.OpenDataSource(fileName)

            ' insert some text and the mailmergefields defined in the datasource
            application.Selection.TypeText("This test is brought to you by ")
            newDocument.MailMerge.Fields.Add(application.Selection.Range, "ProjectName")

            application.Selection.TypeText(" for more information and examples visit ")
            newDocument.MailMerge.Fields.Add(application.Selection.Range, "ProjectLink ")

            application.Selection.TypeText(" or click ")

            newDocument.Hyperlinks.Add(application.Selection.Range, newDocument.MailMerge.DataSource.DataFields(2).Value, _
                                       Missing.Value, "click tooltip", "here", Missing.Value)

            'show the contents of the fields
            newDocument.MailMerge.ViewMailMergeFieldCodes = WdConstants.wdToggle

            'do not show the fieldcodes
            application.ActiveWindow.View.ShowFieldCodes = False

            newDocument.Close(False)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit(WdSaveOptions.wdDoNotSaveChanges)
                application.Dispose()
            End If

        End Try

    End Function

End Class
