Imports System.Reflection

Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication As New Outlook.Application()

        ' we register some events. note: the event trigger was called from word, means an other Thread
        ' remove the Quit() call below and check out more events if you want
 
        Dim mailItem As Outlook.MailItem = outlookApplication.CreateItem(OlItemType.olMailItem)

        Dim closeHandler As Outlook.MailItem_CloseEventHandler = AddressOf Me.mailItem_CloseEvent
        AddHandler mailItem.CloseEvent, closeHandler

        mailItem.BodyFormat = OlBodyFormat.olFormatPlain
        mailItem.Body = "Body of Example06 " + DateTime.Now.ToLongTimeString()
        mailItem.Subject = "Example06"
        mailItem.Display()
        mailItem.Close(OlInspectorClose.olDiscard)
 
        ' close word and dispose reference
        outlookApplication.Quit()
        outlookApplication.Dispose()

    End Sub

    Private Sub mailItem_CloseEvent(ByRef Cancel As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event Close called."})

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

End Class
