Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example01"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Inbox Folder"
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

        ' Get inbox 
        Dim outlookNS As Outlook._NameSpace = outlookApplication.GetNamespace("MAPI")
        Dim inboxFolder As Outlook.MAPIFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

        ' setup gui
        listViewInboxFolder.Items.Clear()
        labelItemsCount.Text = String.Format("You have {0} e-mails.", inboxFolder.Items.Count)

        For Each item As ICOMObject In inboxFolder.Items

            ' not every item in the inbox folder is a mail item
            If (TypeName(item) = "MailItem") Then
                Dim mailItem As Outlook.MailItem = item
                Dim newItem As ListViewItem = listViewInboxFolder.Items.Add(mailItem.SenderName)
                newItem.SubItems.Add(mailItem.Subject)
            End If

        Next

        'close outlook and dispose
        If Not outlookApplication.FromProxyService Then
            outlookApplication.Quit()
        End If
        outlookApplication.Dispose()

    End Sub

End Class
