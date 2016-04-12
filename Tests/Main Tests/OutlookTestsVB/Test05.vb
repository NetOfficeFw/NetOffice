Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports Tests.Core

Public Class Test05
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Fetch Contacts."
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
            Return "Outlook"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Outlook.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.OutlookApi.Application()
            NetOffice.OutlookSecurity.Suppress.Enabled = True

            ' enum contacts 
            Dim contactFolder As Outlook.MAPIFolder = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)

            For index As Integer = 1 To contactFolder.Items.Count

                If (TypeName(contactFolder.Items(index)) = "ContactItem") Then
                    Dim contact As Outlook.ContactItem = contactFolder.Items(index)
                    Console.WriteLine(contact.CompanyAndFullName)
                End If

            Next index

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, String.Format("{0} ContactFolder Items.", contactFolder.Items.Count))

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
