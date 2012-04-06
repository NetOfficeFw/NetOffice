Imports LateBindingApi.Core
Imports Word = NetOffice.WordApi
Imports Office = NetOffice.OfficeApi
Imports NetOffice.WordApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Example07
    Implements IExample

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate
    Dim _wordApplication As Word.Application
    Dim _hostApplication As ExampleBase.IHost

    Public Sub New()

        InitializeComponent()

        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example07", "Beispiel07")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Customize UI", "UI Items erstellen")
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

        Dim commandBar As Office.CommandBar = Nothing
        Dim commandBarBtn As Office.CommandBarButton = Nothing

        ' start word and turn off msg boxes
        _wordApplication = New Word.Application()
        _wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone

        ' add a new document
        _wordApplication.Documents.Add()

        Dim normalDotTemplate As Word.Template = GetNormalDotTemplate()
        _wordApplication.CustomizationContext = normalDotTemplate

        ' add a commandbar popup
        Dim commandBarPopup As Office.CommandBarPopup = _wordApplication.CommandBars("Menu Bar").Controls.Add(MsoControlType.msoControlPopup, MsoBarPosition.msoBarTop, System.Type.Missing, 1, True)
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
        Clipboard.SetDataObject(_hostApplication.DisplayIcon.ToBitmap())
        commandBarBtn.PasteFace()
        Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        'add a new toolbar
        commandBar = _wordApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, False, True)
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
        Clipboard.SetDataObject(_hostApplication.DisplayIcon.ToBitmap())
        commandBarBtn.PasteFace()
        clickHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' create context menu
        commandBarPopup = _wordApplication.CommandBars("Text").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, 1, True)
        commandBarPopup.Caption = "commandBarPopup"

        ' add a button to the popup
        commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        commandBarBtn.FaceId = 9
        clickHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        normalDotTemplate.Saved = True

        ' make visible & set buttons
        _wordApplication.Visible = MsoTriState.msoTrue
        buttonStartExample.Enabled = False
        buttonQuitExample.Enabled = True

    End Sub

    Private Sub buttonQuitExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonQuitExample.Click

        _wordApplication.Quit()
        _wordApplication.Dispose()

        buttonStartExample.Enabled = True
        buttonQuitExample.Enabled = False

    End Sub

#End Region

#Region "Word Trigger"

    Private Sub commandBarBtn_Click(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Click called."})
        Ctrl.Dispose()

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' returns normal.dot template (normal.dotm in modern word versions)
    ''' </summary>
    ''' <remarks></remarks>
    Private Function GetNormalDotTemplate() As Word.Template

        For Each installedTemplate As Word.Template In _wordApplication.Templates
            If (installedTemplate.Name.StartsWith("normal", StringComparison.InvariantCultureIgnoreCase)) Then

                Return installedTemplate
            End If
        Next

        Throw New IndexOutOfRangeException("Template not found.")

    End Function

#End Region

End Class
