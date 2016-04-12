Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports Tests.Core

Public Class Test01
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Fetch inbox folder."
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
            Return "Outlook"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Outlook.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.OutlookApi.Application()
            NetOffice.OutlookSecurity.Suppress.Enabled = True

            ' Get inbox 
            Dim outlookNS As Outlook._NameSpace = application.GetNamespace("MAPI")
            Dim inboxFolder As Outlook.MAPIFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

            Dim items As Outlook._Items = inboxFolder.Items
            Dim item As COMObject = Nothing
            Dim i As Integer = 1

            Do

                If (item Is Nothing) Then
                    item = items.GetFirst()
                End If

                'not every item is a mail item
                If (TypeName(item) = "MailItem") Then
                    Dim mailItem As Outlook.MailItem = item
                    Console.WriteLine(mailItem.SenderName)
                End If

                If Not IsNothing(item) Then
                    item.Dispose()
                End If

                item = items.GetNext()
                i += 1

            Loop While (Not item Is Nothing)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, String.Format("{0} Inbox Items.", items.Count))

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
