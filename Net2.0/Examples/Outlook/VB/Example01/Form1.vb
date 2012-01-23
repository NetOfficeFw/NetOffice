Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' Get inbox 
        Dim outlookNS As Outlook._NameSpace = outlookApplication.GetNamespace("MAPI")
        Dim inboxFolder As Outlook.MAPIFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

        ' setup gui
        listView1.Items.Clear()
        labelItemsCount.Text = String.Format("You have {0} e-mails.", inboxFolder.Items.Count)

        ' we fetch the inbox folder items. ATTENTION: items is null if you have no items in inbox folder
        ' office products initialize ALL collections on demand. this is just an example, we dont check for null here
        ' NOTE: for some uninitialized collections you get an exception while accessing
        For Each item As COMObject In inboxFolder.Items

            'not every item is a mail item
            If (TypeName(item) = "MailItem") Then
                Dim mailItem As Outlook.MailItem = item
                Dim newItem As ListViewItem = listView1.Items.Add(mailItem.SenderName)
                newItem.SubItems.Add(mailItem.Subject)
            End If

            item.Dispose()

        Next

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

    End Sub

End Class
