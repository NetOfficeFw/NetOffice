Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' Create a new MailItem.
        Dim mailItem As Outlook.MailItem = outlookApplication.CreateItem(OlItemType.olMailItem)

        ' prepare item and send
        mailItem.Recipients.Add(textBoxReciever.Text)
        mailItem.Subject = textBoxSubject.Text
        mailItem.Body = textBoxBody.Text
        mailItem.Send()

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class
