Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports Tests.Core

Public Class Test06
    Implements ITestPackage

    Private _closeEventCalled As Boolean

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Test events."
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
            Return "Outlook"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Outlook.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try

            application = New NetOffice.OutlookApi.Application()
            NetOffice.OutlookSecurity.Suppress.Enabled = True

            Dim mailItem As Outlook.MailItem = application.CreateItem(OlItemType.olMailItem)

            Dim closeHandler As Outlook.MailItem_CloseEventHandler = AddressOf Me.mailItem_CloseEvent
            AddHandler mailItem.CloseEvent, closeHandler

            ' BodyFormat is not available in Outlook 2000
            ' we check at runtime is property is available
            If (mailItem.EntityIsAvailable("BodyFormat")) Then
                mailItem.BodyFormat = OlBodyFormat.olFormatPlain
            End If
            mailItem.Body = "NetOffice VB Test06" + DateTime.Now.ToLongTimeString()
            mailItem.Subject = "Test06"
            mailItem.Display()
            mailItem.Close(OlInspectorClose.olDiscard)

            If (_closeEventCalled) Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Return New TestResult(False, DateTime.Now.Subtract(startTime), "CloseEvent not triggered.", Nothing, "")
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


    Private Sub mailItem_CloseEvent(ByRef Cancel As Boolean)

        _closeEventCalled = True

    End Sub

End Class
