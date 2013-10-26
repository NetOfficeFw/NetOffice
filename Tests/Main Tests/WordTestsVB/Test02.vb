Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test02
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using a DataTable."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test02"
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

            ' add a table
            Dim table As Word.Table
            table = newDocument.Tables.Add(application.Selection.Range, 3, 2)

            'insert some text into the cells
            table.Cell(1, 1).Select()
            application.Selection.TypeText("This")

            table.Cell(1, 2).Select()
            application.Selection.TypeText("table")

            table.Cell(2, 1).Select()
            application.Selection.TypeText("was")

            table.Cell(2, 2).Select()
            application.Selection.TypeText("created")

            table.Cell(3, 1).Select()
            application.Selection.TypeText("by")

            table.Cell(3, 2).Select()
            application.Selection.TypeText("NetOffice")

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
