Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Tests.Core

Public Class Test08
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using Paragraphes."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test08"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Word"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim document As Word.Document = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try

            document = New Word.Document()
            document.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone
            document.Application.Selection.TypeText("Test with TabIntend VB")
            document.Application.Selection.Start = 0
            Dim p As Word.Paragraph = document.Application.Selection.Range.Paragraphs(1)

            p.IndentCharWidth(10)
            p.IndentFirstLineCharWidth(8)
            p.Space1()
            p.Space15()
            p.Space2()
            p.TabHangingIndent(5)
            p.TabIndent(3)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(document) Then
                document.Application.Quit()
                document.Dispose()
            End If

        End Try

    End Function

End Class
