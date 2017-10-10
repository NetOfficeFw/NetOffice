Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example03
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example03"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Send an E- Mail"
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

        ' Create a new MailItem.
        Dim mailItem As Outlook.MailItem = outlookApplication.CreateItem(OlItemType.olMailItem)

        ' prepare item and send
        mailItem.Recipients.Add(textBoxReciever.Text)
        mailItem.Subject = textBoxSubject.Text
        mailItem.Body = textBoxBody.Text
        mailItem.Send()

        'close outlook and dispose
        If Not outlookApplication.FromProxyService Then
            outlookApplication.Quit()
        End If
        outlookApplication.Dispose()

        _hostApplication.ShowFinishDialog("Done!", Nothing)

    End Sub

End Class
