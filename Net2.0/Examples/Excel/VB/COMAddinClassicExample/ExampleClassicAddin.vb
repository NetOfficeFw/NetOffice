Imports System.Reflection
Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<GuidAttribute("E4D04E40-5759-4cec-9868-FE475C051DC8"), ProgIdAttribute("ExcelClassicAddinVB.Addin")> _
Public Class ExampleClassicAddin
    Implements IDTExtensibility2

    Private Shared ReadOnly _addinRegistryKey As String = "Software\\Microsoft\\Office\\Excel\\AddIns\\"
    Private Shared ReadOnly _prodId As String = "ExcelClassicAddinVB.Addin"
    Private Shared ReadOnly _addinName As String = "VB ExampleClassicAddin"

    ' gui elements
    Private Shared ReadOnly _toolbarName = "VB_COMAddinClassicToolbar"
    Private Shared ReadOnly _toolbarButtonName As String = "VB_ToolbarButton"
    Private Shared ReadOnly _toolbarPopupName As String = "VB_COMAddinClassicPopup"
    Private Shared ReadOnly _menuName As String = "VB_COMAddinClassicMenu"
    Private Shared ReadOnly _menuButtonName As String = "VB_MenuButton"
    Private Shared ReadOnly _contextName As String = "VB_COMAddinClassicContext"
    Private Shared ReadOnly _contextMenuButtonName As String = "VB_ContextButton"

    Dim _excelApplication As Excel.Application = Nothing

#Region "IDTExtensibility2 Members"

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        Try

            ' Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize()

            _excelApplication = New Excel.Application(Nothing, Application)

        Catch ex As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            ' we save the error info in addin registry key

            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, True)

            rk.SetValue("LastError", "An error occured in OnConnection.")
            rk.SetValue("LastException", ex.Message)
            rk.Close()

        End Try
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection
        Try

            If (Not IsNothing(_excelApplication)) Then
                _excelApplication.Dispose()
            End If

        Catch ex As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            ' we save the error info in addin registry key

            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, True)

            rk.SetValue("LastError", "An error occured in OnDisconnection.")
            rk.SetValue("LastException", ex.Message)
            rk.Close()

        End Try
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete
        Try

            SetupGui()

        Catch ex As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            ' we save the error info in addin registry key

            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, True)

            rk.SetValue("LastError", "An error occured in OnStartupComplete.")
            rk.SetValue("LastException", ex.Message)
            rk.Close()

        End Try
    End Sub

#End Region

#Region "COM Register Functions"

    <ComRegisterFunctionAttribute()> _
    Public Shared Sub RegisterFunction(ByVal type As Type)
        Try

            ' add codebase value
            Dim thisAssembly As Assembly = Assembly.GetAssembly(GetType(ExampleClassicAddin))
            Dim key As RegistryKey = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\1.0.0.0")
            key.SetValue("CodeBase", thisAssembly.CodeBase)
            key.Close()

            key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32")
            key.SetValue("CodeBase", thisAssembly.CodeBase)
            key.Close()

            ' add bypass key
            ' http://support.microsoft.com/kb/948461
            key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}")
            Dim defaultValue As String = key.GetValue("")
            If (IsNothing(defaultValue)) Then
                key.SetValue("", "Office .NET Framework Lockback Bypass Key")
            End If
            key.Close()

            ' add excel addin key
            Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable")
            Registry.CurrentUser.CreateSubKey(_addinRegistryKey + _prodId)
            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, True)
            rk.SetValue("LoadBehavior", CInt(3))
            rk.SetValue("FriendlyName", _addinName)
            rk.SetValue("Description", "NetOffice COMAddinExample with classic UI")
            rk.Close()

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)
            Registry.CurrentUser.DeleteSubKey(_addinRegistryKey + _prodId)

        Catch ex As ArgumentException
            ' key is missing
        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region

#Region "Setup GUI"

    ''' <summary>
    ''' creates gui elements
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupGui()

        ' How to: Add Commands to Shortcut Menus in Excel
        ' http://msdn.microsoft.com/en-us/library/0batekf4.aspx   

        'create commandbar 
        Dim commandBar As Office.CommandBar = _excelApplication.CommandBars.Add(_toolbarName, MsoBarPosition.msoBarTop, System.Type.Missing, True)
        commandBar.Visible = True

        ' add popup to commandbar
        Dim commandBarPop As Office.CommandBarPopup = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _toolbarPopupName
        commandBarPop.Tag = _toolbarPopupName

        'add a button to the popup
        Dim commandBarBtn As Office.CommandBarButton = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.FaceId = 9
        commandBarBtn.Caption = _toolbarButtonName
        commandBarBtn.Tag = _toolbarButtonName
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' create menu 
        commandBar = _excelApplication.CommandBars("Worksheet Menu Bar")

        ' add popup to menu bar
        commandBarPop = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _menuName
        commandBarPop.Tag = _menuName

        ' add a button to the popup
        commandBarBtn = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.FaceId = 9
        commandBarBtn.Caption = _menuButtonName
        commandBarBtn.Tag = _menuButtonName
        clickHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' create context menu 
        commandBarPop = _excelApplication.CommandBars("Cell").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _contextName
        commandBarPop.Tag = _contextName

        ' add a button to the popup
        commandBarBtn = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _contextMenuButtonName
        commandBarBtn.Tag = _contextMenuButtonName
        commandBarBtn.FaceId = 9
        clickHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

    End Sub

#End Region

#Region "Trigger"

    ''' <summary>
    ''' Click event trigger from created buttons. incoming call comes from word application thread.
    ''' </summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    ''' <remarks></remarks>
    Private Sub commandBarBtn_ClickEvent(ByVal Ctrl As NetOffice.OfficeApi.CommandBarButton, ByRef CancelDefault As Boolean)

        Dim message As String = String.Format("Click from Button {0}.", Ctrl.Caption)
        MessageBox.Show(message, _addinName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Ctrl.Dispose()

    End Sub

#End Region

End Class
