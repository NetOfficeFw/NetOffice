Imports ExampleBase
Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

Public Class Example09
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost
    Dim _excelApplication As Excel.Application
    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

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
            Return IIf(_hostApplication.LCID = 1033, "Example09", "Beispiel09")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Customize classic UI without ribbons and recieve click events", "Erweitern der klassischen Oberfläche und beziehen von Click Events")
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

        ' start excel and turn off msg boxes
        _excelApplication = New Excel.Application()
        _excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = _excelApplication.Workbooks.Add()

        ' add a commandbar popup
        Dim commandBarPopup As Office.CommandBarPopup = _excelApplication.CommandBars("Worksheet Menu Bar").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
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
        Clipboard.SetDataObject(_hostApplication.DisplayIcon.ToBitmap())
        commandBarBtn.PasteFace()
        Dim clickHandler As Office.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        'add a new toolbar
        commandBar = _excelApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, System.Type.Missing, True)
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
        commandBarPopup = _excelApplication.CommandBars("Cell").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPopup.Caption = "commandBarPopup"

        ' add a button to the popup
        commandBarBtn = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = "commandBarButton"
        commandBarBtn.FaceId = 9
        clickHandler = AddressOf Me.commandBarBtn_Click
        AddHandler commandBarBtn.ClickEvent, clickHandler

        Dim sheet As Excel.Worksheet = workBook.Worksheets(1)
        sheet.Cells(2, 2).Value = "this excel instance contains 3 custom menus"
        sheet.Cells(3, 2).Value = "the main menu, the toolbar menu and the cell context menu"
        sheet.Cells(4, 2).Value = "in this case the menus are temporaily created"
        sheet.Cells(5, 2).Value = "they are not persistant and needs no unload event or something like this"
        sheet.Cells(6, 2).Value = "you can also create persistant menus if you want"

        ' make visible & set buttons
        _excelApplication.Visible = True
        buttonStartExample.Enabled = False
        buttonQuitExample.Enabled = True

    End Sub

    Private Sub buttonQuitExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonQuitExample.Click

        _excelApplication.Quit()
        _excelApplication.Dispose()

        buttonStartExample.Enabled = True
        buttonQuitExample.Enabled = False

    End Sub

#End Region

#Region "Excel Trigger"

    Private Sub commandBarBtn_Click(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Click called."})
        Ctrl.Dispose()

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

#End Region

End Class
