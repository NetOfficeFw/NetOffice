Imports System.Reflection

Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example06
    Implements IExample

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate
     
    Dim _hostApplication As ExampleBase.IHost

    Public Sub New()

        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example06", "Beispiel06")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Events", "Ereignisse")
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
        Dim outlookApplication As New Outlook.Application()

        ' we register some events. note: the event trigger was called from word, means an other Thread
        Dim mailItem As Outlook.MailItem = outlookApplication.CreateItem(OlItemType.olMailItem)

        Dim closeHandler As Outlook.MailItem_CloseEventHandler = AddressOf Me.mailItem_CloseEvent
        AddHandler mailItem.CloseEvent, closeHandler

        ' BodyFormat is not available in Outlook 2000 we check at runtime is property is available
        If (mailItem.EntityIsAvailable("BodyFormat")) Then
            mailItem.BodyFormat = OlBodyFormat.olFormatPlain
        End If
        mailItem.Body = "ExampleBody"
        mailItem.Subject = "ExampleSubject"
        mailItem.Display()
        mailItem.Close(OlInspectorClose.olDiscard)

        ' close word and dispose reference
        outlookApplication.Quit()
        outlookApplication.Dispose()


    End Sub

#End Region

#Region "Outlook Trigger"

    Private Sub mailItem_CloseEvent(ByRef Cancel As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event Close called."})

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

#End Region

End Class
