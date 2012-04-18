Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports Office = NetOffice.OfficeApi
Imports Tests.Core

Public Class Test09
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Check for loaded Addin"
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test09"
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
            application.Visible = True
            application.DisplayAlerts = False
            application.Workbooks.Add()

            Dim addIn As Office.COMAddIn = Nothing

            For Each item As Office.COMAddIn In application.COMAddIns

                If item.ProgId = "ExcelAddinVB.TestAddin" Then
                    addIn = item
                    Exit For
                End If

            Next

            If IsNothing(addIn) Then
                Return New TestResult(False, DateTime.Now.Subtract(startTime), "COMAddin ExcelAddinVB.TestAddin not found.", Nothing, "")
            End If

            Dim addinProxy As COMObject = New COMObject(Nothing, addIn.Object)
            Dim ribbonIsOkay As Boolean = Invoker.PropertyGet(addinProxy, "RibbonUIPassed")
            Dim taskPaneIsOkay As Boolean = Invoker.PropertyGet(addinProxy, "TaskPanePassed")
            addinProxy.Dispose()

            If ribbonIsOkay = True And taskPaneIsOkay = True Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Return New TestResult(False, DateTime.Now.Subtract(startTime), String.Format("Ribbon:{0} TaskPane:{1}", ribbonIsOkay, taskPaneIsOkay), Nothing, "")
            End If

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
