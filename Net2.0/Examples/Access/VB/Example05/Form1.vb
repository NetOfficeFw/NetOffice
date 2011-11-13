Imports System.Reflection

Imports LateBindingApi.Core

Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports NetOffice.AccessApi.Constants

Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

Imports DAO = NetOffice.DAOApi
Imports NetOffice.DAOApi.Enums
Imports NetOffice.DAOApi.Constants

Public Class Form1

    Dim _accessApplication As Access.Application
    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        Dim commandBar As Office.CommandBar = Nothing
        Dim commandBarBtn As Office.CommandBarButton = Nothing

        ' start access 
        _accessApplication = New Access.Application()

        ' add a commandbar popup
        Dim commandBarPopup As Office.CommandBarPopup = _accessApplication.CommandBars("Menu Bar").Controls.Add(MsoControlType.msoControlPopup)
        commandBarPopup.Caption = "commandBarPopup"

        ' you can see we use an own icon via .PasteFace()
        ' is not possible from outside process boundaries to use the PictureProperty directly
        ' the reason for is IPictureDisp: http://support.microsoft.com/kb/286460/de
        ' its not important is early or late binding or managed or unmanaged, the behaviour is always the same
        ' For example, a COMAddin running as InProcServer and can access the Picture Property

        ' add a button to the popup
        commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        Clipboard.SetDataObject(Me.Icon.ToBitmap())
        commandBarBtn.PasteFace()
        Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' make visible & set buttons
        _accessApplication.Visible = MsoTriState.msoTrue
        button1.Enabled = False
        button2.Enabled = True

    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click

        _accessApplication.Quit(AcQuitOption.acQuitSaveNone)
        _accessApplication.Dispose()

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
