Imports System.Reflection
Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports LateBindingApi.Core
Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi
Imports Word = NetOffice.WordApi
Imports Outlook = NetOffice.OutlookApi
Imports PowerPoint = NetOffice.PowerPointApi
Imports Access = NetOffice.AccessApi

Imports Extensibility


<GuidAttribute("76345C06-E899-4762-8F75-C54F49D813B9"), ProgIdAttribute("SuperAddinVB.Connect")> _
Public Class Connect
    Implements IDTExtensibility2, IRibbonExtensibility

    Private Shared ReadOnly _prodId As String = "SuperAddinVB.Connect"
    Private Shared ReadOnly _addinName As String = "SuperAddinVB"

#Region "Fields"

    Private _application As HostApplication
    Private _trayIcon As TrayIcon

    Private _isRibbonSupported As Boolean

#End Region

#Region "IDTExtensibility2 Members"

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown

    End Sub

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As Extensibility.ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        Try

            ' initialize api
            LateBindingApi.Core.Factory.Initialize()

            _application = New HostApplication(Application, ConnectMode, AddInInst, custom)
            _trayIcon = New TrayIcon(True)

        Catch throwedException As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            FormShowError.LogError("An error ocurred while perform OnConnection.", throwedException)

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection

        Try

            If (Not IsNothing(_application)) Then
                _application.Dispose()
            End If

            If (Not IsNothing(_trayIcon)) Then
                _trayIcon.Dispose()
            End If


        Catch throwedException As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            FormShowError.LogError("An error ocurred while perform OnDisconnection.", throwedException)

        End Try

    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete

        Try

            If (False = _isRibbonSupported) Then
                CreateClassicUI()
            End If

        Catch throwedException As Exception

            ' dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
            FormShowError.LogError("An error ocurred while perform OnStartupComplete.", throwedException)

        End Try

    End Sub

#End Region

#Region "COM Register Functions"

    <ComRegisterFunctionAttribute()> _
    Public Shared Sub RegisterFunction(ByVal type As Type)

        Try

            ' add codebase value
            Dim thisAssembly As Assembly = Assembly.GetAssembly(GetType(Connect))
            Dim key As RegistryKey = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0")
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

            Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable")

            OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Excel + _prodId)
            OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Word + _prodId)
            OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Outlook + _prodId)
            OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.PowerPoint + _prodId)
            OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Access + _prodId)

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)

            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Excel + _prodId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Word + _prodId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Outlook + _prodId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.PowerPoint + _prodId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Access + _prodId)

        Catch ex As ArgumentException
            ' key is missing
        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI

        Try

            _isRibbonSupported = True
            Return ReadString("RibbonUI.xml")

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured in GetCustomUI." + details, "GetCustomUI " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""

        End Try

    End Function

    Public Sub OnAction(ByVal control As IRibbonControl)

        Try

            Dim messageString As String = String.Format("Thanks for click on a Ribbon." + vbNewLine + "HostApp is {0}.{1} Version:{2}", _
                                             _application.ComponentName, _application.Name, _application.Version)

            MessageBox.Show(messageString, "SuperAddin", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#Region "Classic UI"

    Private Sub commandBarBtn_ClickEvent(ByVal Ctrl As Office.CommandBarButton, ByRef CancelDefault As Boolean)

        Dim message As String = String.Format("Thanks for click on a button." + vbNewLine + "HostApp is {0}.{1} Version:{2}", _
                                                _application.ComponentName, _application.Name, _application.Version)
        MessageBox.Show(message, _addinName, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Ctrl.Dispose()

    End Sub

    ''' <summary>
    ''' calls specific create method for office application type
    ''' note: all applications has the same code except outlook (active inspector)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateClassicUI()

        If _application.ComponentName = "Microsoft Excel" Or _application.ComponentName = "Excel" Then
            CreateExcelUI()
        ElseIf _application.ComponentName = "Microsoft Word" Or _application.ComponentName = "Word" Then
            CreateWordUI()
        ElseIf _application.ComponentName = "Microsoft Outlook" Or _application.ComponentName = "Outlook" Then
            CreateOutlookUI()
        ElseIf _application.ComponentName = "Microsoft PowerPoint" Or _application.ComponentName = "PowerPoint" Then
            CreatePowerPointUI()
        ElseIf _application.ComponentName = "Microsoft Access" Or _application.ComponentName = "Access" Then
            CreateAccessUI()
        End If

    End Sub

    Private Sub CreateExcelUI()

        Dim excelApp As Excel.Application = _application.Application

        Dim commandBar As Office.CommandBar = excelApp.CommandBars.Add(_addinName + "Commandbar", Office.Enums.MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        ' add a button to the toolbar
        Dim commandBarBtn As Office.CommandBarButton = commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, True)
        commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _addinName + "ExcelButton"
        commandBarBtn.FaceId = 2
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        excelApp.DisposeChildInstances(False)

    End Sub

    Private Sub CreateOutlookUI()

        Dim outlookApp As Outlook.Application = _application.Application
        Dim commandBar As Office.CommandBar = outlookApp.ActiveExplorer().CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        'add a button to the toolbar
        Dim commandBarBtn As Office.CommandBarButton = commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, True)
        commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _addinName + "OutlookButton"
        commandBarBtn.FaceId = 2
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        outlookApp.DisposeChildInstances(False)

    End Sub

    Private Sub CreateWordUI()

        Dim wordApp As Word.Application = _application.Application
        Dim commandBar As Office.CommandBar = wordApp.CommandBars.Add(_addinName + "Commandbar", Office.Enums.MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        ' add a button to the toolbar
        Dim commandBarBtn As Office.CommandBarButton = commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, True)
        commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _addinName + "WordButton"
        commandBarBtn.FaceId = 2
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        wordApp.DisposeChildInstances(False)

    End Sub

    Private Sub CreatePowerPointUI()

        Dim powerApp As PowerPoint.Application = _application.Application

        Dim commandBar As Office.CommandBar = powerApp.CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        ' add a button to the toolbar
        Dim commandBarBtn As Office.CommandBarButton = commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, True)
        commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _addinName + "PowerButton"
        commandBarBtn.FaceId = 2
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        powerApp.DisposeChildInstances(False)

    End Sub

    Private Sub CreateAccessUI()

        Dim accessApp As Access.Application = _application.Application

        Dim commandBar As Office.CommandBar = accessApp.CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, False, True)
        commandBar.Visible = True

        ' add a button to the toolbar
        Dim commandBarBtn As Office.CommandBarButton = commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, True)
        commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _addinName + "AccessButton"
        commandBarBtn.FaceId = 2
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        accessApp.DisposeChildInstances(False)

    End Sub

#End Region

#Region "Private Helper"

    ''' <summary>
    ''' reads text from ressource
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function ReadString(ByVal fileName As String) As String

        fileName = _addinName + "." + fileName

        Dim ressourceStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName)
        If (IsNothing(ressourceStream)) Then
            Throw (New System.IO.IOException("Error accessing resource Stream."))
        End If

        Dim textStreamReader As System.IO.StreamReader = New System.IO.StreamReader(ressourceStream)
        If (IsNothing(textStreamReader)) Then
            Throw (New System.IO.IOException("Error accessing resource File."))
        End If

        Dim text As String = textStreamReader.ReadToEnd()
        ressourceStream.Close()
        textStreamReader.Close()
        Return text

    End Function

#End Region

End Class
