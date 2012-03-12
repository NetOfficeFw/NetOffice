Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
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
                Dim listItem As ListViewItem = listView1.Items.Add(index.ToString())
                listItem.SubItems.Add(contact.CompanyAndFullName)
            End If

        Next
         
        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class
