Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports Tests.Core

Public Class Test03
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Send a mail."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test03"
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

            ' Create a new MailItem.
            Dim mailItem As Outlook.MailItem = application.CreateItem(OlItemType.olMailItem)

            ' prepare item and send
            mailItem.Recipients.Add("public.sebastian@web.de")
            mailItem.Subject = "NetOffice Test Mail(VB)"
            mailItem.Body = "This is a NetOffice test mail from the MainTests.(VB)"
            mailItem.Send()

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
