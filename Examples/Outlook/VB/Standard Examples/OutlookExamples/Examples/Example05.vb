Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example05
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example05"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "List all contacts"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Me
        End Get
    End Property

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' start outlook by trying to access running application first
        Dim outlookApplication = New Outlook.Application(True)

        ' enum contacts 
        Dim index As Integer
        Dim contactFolder As Outlook.MAPIFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)

        For Each item As ICOMObject In contactFolder.Items

            If (TypeName(item) = "ContactItem") Then
                index += 1
                Dim contact As Outlook.ContactItem = item
                Dim listItem As ListViewItem = listViewContacts.Items.Add(index.ToString())
                listItem.SubItems.Add(contact.CompanyAndFullName)
            End If

        Next


        'close outlook and dispose
        If Not outlookApplication.FromProxyService Then
            outlookApplication.Quit()
        End If
        outlookApplication.Dispose()

    End Sub


End Class
