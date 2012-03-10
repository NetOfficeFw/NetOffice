Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' SendAndReceive is supported from Outlooks 2007 or higher
        ' we check at runtime the feature is available
        If outlookApplication.Session.EntityIsAvailable("SendAndReceive") Then
            ' one simple call
            outlookApplication.Session.SendAndReceive(False)
        Else
            MessageBox.Show(Me, "This version of MS-Outlook doesnt support SendAndReceive.", "Example04", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
       
        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class
