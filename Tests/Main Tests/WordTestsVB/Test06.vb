Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test06
    Implements ITestPackage

    Dim _newDocumentCalled As Boolean
    Dim _documentBeforeCloseCalled As Boolean

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using events."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test06"
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

            ' we register some events. note: the event trigger was called from word, means an other Thread
            ' remove the Quit() call below and check out more events if you want

            Dim newHandler As Word.Application_NewDocumentEventHandler = AddressOf Me.wordApplication_NewDocumentEvent
            AddHandler application.NewDocumentEvent, newHandler

            Dim newCloseHandler As Word.Application_DocumentBeforeCloseEventHandler = AddressOf Me.wordApplication_DocumentBeforeCloseEvent
            AddHandler application.DocumentBeforeCloseEvent, newCloseHandler

            ' add a document and close
            Dim document As Word.Document = application.Documents.Add()
            document.Close()

            If (_documentBeforeCloseCalled And _newDocumentCalled) Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Return New TestResult(False, DateTime.Now.Subtract(startTime), String.Format("DocumentBeforeClose:{0}, NewDocument:{1}", _documentBeforeCloseCalled, _newDocumentCalled), Nothing, "")
            End If

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit(WdSaveOptions.wdDoNotSaveChanges)
                application.Dispose()
            End If

        End Try

    End Function
     
    Private Sub wordApplication_NewDocumentEvent(ByVal Doc As Word.Document)

        _newDocumentCalled = True
        Doc.Dispose()

    End Sub


    Private Sub wordApplication_DocumentBeforeCloseEvent(ByVal Doc As Word.Document, ByRef Cancel As Boolean)

        _documentBeforeCloseCalled = True
        Doc.Dispose()

    End Sub

End Class
