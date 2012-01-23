Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' enum contacts 
        Dim i As Integer
        Dim contactFolder As Outlook.MAPIFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)
        For Each item As COMObject In contactFolder.Items

            If (TypeName(item) = "ContactItem") Then

                i = i + 1
                Dim contactItem As Outlook.ContactItem = item
                Dim listItem As ListViewItem = listView1.Items.Add(i.ToString())
                listItem.SubItems.Add(contactItem.CompanyAndFullName)

            End If

        Next

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class
