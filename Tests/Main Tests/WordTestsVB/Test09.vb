Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
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
            Return "Word"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Word.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New Word.Application()
            application.Visible = True
            application.DisplayAlerts = NetOffice.WordApi.Enums.WdAlertLevel.wdAlertsNone
            application.Documents.Add()

            Dim addIn As Office.COMAddIn = Nothing

            For Each item As Office.COMAddIn In application.COMAddIns

                If item.ProgId = "NOTestsMain.WordTestAddinVB" Then
                    addIn = item
                    Exit For
                End If

            Next

            If (IsNothing(addIn) Or IsNothing(addIn.Object)) Then
                Return New TestResult(False, DateTime.Now.Subtract(startTime), "COMAddin NOTestsMain.WordTestAddinVB or addIn.Object not found.", Nothing, "")
            End If

            Dim addinProxy As COMObject = New COMObject(addIn.Object)
            Dim addinStatusOkay As Boolean = Invoker.Default.PropertyGet(addinProxy, "StatusOkay")
            Dim addinStatusDescription As String = Invoker.Default.PropertyGet(addinProxy, "StatusDescription")
            addinProxy.Dispose()

            If addinStatusOkay = True Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Return New TestResult(False, DateTime.Now.Subtract(startTime), String.Format("NOTestsMain.WordTestAddinVB Addin Status {0}", addinStatusDescription), Nothing, "")
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

End Class
