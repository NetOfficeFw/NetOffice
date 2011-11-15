Imports System.Reflection

Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Dim _outlookApplication As Outlook.Application
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

        Dim commandBar As Office.CommandBar = Nothing
        Dim commandBarBtn As Office.CommandBarButton = Nothing

        ' start outlook
        _outlookApplication = New Outlook.Application()

        Dim outlookNS As Outlook._NameSpace = _outlookApplication.GetNamespace("MAPI")
        Dim inboxFolder As Outlook.MAPIFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
        inboxFolder.Display()

        ' add a commandbar popup
        Dim commandBarPopup As Office.CommandBarPopup = _outlookApplication.ActiveExplorer().CommandBars("Menu Bar").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPopup.Caption = "commandBarPopup"

        ' you can see we use an own icon via .PasteFace()
        ' is not possible from outside process boundaries to use the PictureProperty directly
        ' the reason for is IPictureDisp: http://support.microsoft.com/kb/286460/de
        ' its not important is early or late binding or managed or unmanaged, the behaviour is always the same
        ' For example, a COMAddin running as InProcServer and can access the Picture Property

        ' add a button to the popup
        commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        Clipboard.SetDataObject(Me.Icon.ToBitmap())
        commandBarBtn.PasteFace()
        Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        'add a new toolbar
        commandBar = _outlookApplication.ActiveExplorer().CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        ' add a button to the toolbar
        commandBarBtn = commandBar.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        commandBarBtn.FaceId = 3
        clickHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' add a dropdown box to the toolbar
        commandBarPopup = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPopup.Caption = "commandBarPopup"

        ' add a button to the popup, we use an own icon for the button
        commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        Clipboard.SetDataObject(Me.Icon.ToBitmap())
        commandBarBtn.PasteFace()
        clickHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        '  set buttons
        button1.Enabled = False
        button2.Enabled = True

    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click

        _outlookApplication.Quit()
        _outlookApplication.Dispose()

        button1.Enabled = True
        button2.Enabled = False

    End Sub

    Private Sub commandBarBtn_Click(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Click called."})
        Ctrl.Dispose()

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

End Class
