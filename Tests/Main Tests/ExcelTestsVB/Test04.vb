Imports Excel = NetOffice.ExcelApi
Imports NetOffice.OfficeApi.Enums
Imports Tests.Core

Public Class Test04
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using shapes."
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
            Return "Excel"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Excel.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False
            application.Workbooks.Add()

            Dim workSheet As Excel.Worksheet = application.Workbooks(1).Sheets(1)

            workSheet.Cells(1, 1).Value = "these sample shapes was dynamicly created by code."

            ' create a star
            Dim starShape As Excel.Shape = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20)

            ' create a simple textbox
            Dim textBox As Excel.Shape = workSheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50)
            textBox.TextFrame.Characters().Text = "text"
            textBox.TextFrame.Characters().Font.Size = 14

            'create a wordart
            Dim textEffect As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12, _
                                                                                MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250)

            ' create text effect
            Dim textDiagram As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14, _
                                                                                MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350)


            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function
End Class
