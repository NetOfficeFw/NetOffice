Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example05
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example05", "Beispiel05")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "List all contacts", "Alle Kontakte auflisten")
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

#End Region

#Region "UI Trigger"

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' enum contacts 
        Dim index As Integer
        Dim contactFolder As Outlook.MAPIFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)

        For Each item As COMObject In contactFolder.Items

            If (TypeName(item) = "ContactItem") Then
                index += 1
                Dim contact As Outlook.ContactItem = contactFolder.Items(index)
                Dim listItem As ListViewItem = listViewContacts.Items.Add(index.ToString())
                listItem.SubItems.Add(contact.CompanyAndFullName)
            End If

        Next

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

    End Sub

#End Region

End Class
