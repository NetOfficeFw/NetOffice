Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example04
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' SendAndReceive is supported from Outlooks 2007 or higher. we check at runtime the feature is available
        If outlookApplication.Session.EntityIsAvailable("SendAndReceive") Then
            ' one simple call
            outlookApplication.Session.SendAndReceive(False)
        Else
            _hostApplication.ShowErrorDialog("This version of MS-Outlook doesnt support SendAndReceive.", Nothing)
        End If

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        _hostApplication.ShowFinishDialog("Done!", Nothing)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example04", "Beispiel04")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Send and Recieve", "Senden und empfangen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

End Class
