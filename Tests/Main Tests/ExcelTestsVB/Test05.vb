Imports Excel = NetOffice.ExcelApi
Imports Tests.Core

Public Class Test05
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using charts and datasource."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test05"
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

            ' we need some data to display
            Dim dataRange As Excel.Range = PutSampleData(workSheet)

            ' create a nice diagram
            Dim chartObjects As Excel.ChartObjects = workSheet.ChartObjects()
            Dim chart As Excel.ChartObject = chartObjects.Add(70, 100, 375, 225)
            chart.Chart.SetSourceData(dataRange)

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

    Private Function PutSampleData(ByVal workSheet As Excel.Worksheet) As Excel.Range

        workSheet.Cells(2, 2).Value = "Datum"
        workSheet.Cells(3, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(4, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(5, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(6, 2).Value = DateTime.Now.ToShortDateString()

        workSheet.Cells(2, 3).Value = "Columns1"
        workSheet.Cells(3, 3).Value = 25
        workSheet.Cells(4, 3).Value = 33
        workSheet.Cells(5, 3).Value = 30
        workSheet.Cells(6, 3).Value = 22

        workSheet.Cells(2, 4).Value = "Column2"
        workSheet.Cells(3, 4).Value = 25
        workSheet.Cells(4, 4).Value = 33
        workSheet.Cells(5, 4).Value = 30
        workSheet.Cells(6, 4).Value = 22

        workSheet.Cells(2, 5).Value = "Column3"
        workSheet.Cells(3, 5).Value = 25
        workSheet.Cells(4, 5).Value = 33
        workSheet.Cells(5, 5).Value = 30
        workSheet.Cells(6, 5).Value = 22

        Return workSheet.get_Range("$B2:$E6")

    End Function

End Class
