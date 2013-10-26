Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test01
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Insert text."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test01"
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

            ' insert some text
            application.Selection.TypeText("This text is written by NetOffice")

            application.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend)
            application.Selection.Font.Color = WdColor.wdColorSeaGreen
            application.Selection.Font.Bold = 1
            application.Selection.Font.Size = 18
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
