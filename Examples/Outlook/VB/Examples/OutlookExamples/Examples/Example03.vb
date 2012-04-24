Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example03
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example03", "Beispiel03")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Send an E- Mail", "Eine E-Mail verschicken")
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

        _hostApplication.ShowFinishDialog("Done!", Nothing)

    End Sub

#End Region

End Class
