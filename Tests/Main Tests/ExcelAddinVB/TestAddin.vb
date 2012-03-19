Imports System.Reflection
Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<GuidAttribute("9779BAB6-021F-4D3A-834F-6E4634CD1EDB"), ProgIdAttribute("ExcelAddinVB.TestAddin")> _
Public Class TestAddin
    Implements IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer

    Private Shared ReadOnly _addinRegistryKey As String = "Software\\Microsoft\\Office\\Excel\\AddIns\\"
    Private Shared ReadOnly _prodId As String = "ExcelAddinVB.TestAddin"
    Private Shared ReadOnly _addinName As String = "TestAddin VB Excel"

    Dim _excelApplication As Excel.Application = Nothing

    Dim _ribbonUIPassed As Boolean
    Dim _taskPanePassed As Boolean

    Public ReadOnly Property RibbonUIPassed As Boolean
        Get
            Return _ribbonUIPassed
        End Get
    End Property

    Public ReadOnly Property TaskPanePassed As Boolean
        Get
            Return _taskPanePassed
        End Get
    End Property

#Region "ICustomTaskPaneConsumer Member"

    Private _myCtpFactory As Office.ICTPFactory
    Private _myPane As Office._CustomTaskPane
    Private _myControl As SampleControl

    Public Sub CTPFactoryAvailable(ByVal CTPFactoryInst As Object) Implements ICustomTaskPaneConsumer.CTPFactoryAvailable

        Try
            _myCtpFactory = New NetOffice.OfficeApi.ICTPFactory(_excelApplication, CTPFactoryInst)
            _myPane = _myCtpFactory.CreateCTP("ExcelAddinVB.SampleControl", "NetOffice Sample Pane (VB)", Type.Missing)
            _myPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft
            _myPane.Visible = True
            _myControl = _myPane.ContentControl
            _taskPanePassed = True
        Catch ex As Exception
            _taskPanePassed = False
            Console.WriteLine(ex.Message)
        End Try

    End Sub

#End Region

#Region "IDTExtensibility2 Members"

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        Try

            ' initialize api
            LateBindingApi.Core.Factory.Initialize()

            _excelApplication = New Excel.Application(Nothing, Application)

            Dim addin As New Office.COMAddIn(Nothing, AddInInst)
            addin.Object = Me

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

    End Sub

#End Region

#Region "COM Register Functions"

    <ComRegisterFunctionAttribute()> _
    Public Shared Sub RegisterFunction(ByVal type As Type)
        Try
            ' add codebase value
            Dim thisAssembly As Assembly = Assembly.GetAssembly(GetType(TestAddin))
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

            ' add excel addin key
            Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable")
            Registry.CurrentUser.CreateSubKey(_addinRegistryKey + _prodId)
            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, True)
            rk.SetValue("LoadBehavior", CInt(3))
            rk.SetValue("FriendlyName", _addinName)
            rk.SetValue("Description", "NetOffice COMAddinExample with ribbon UI")
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

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI

        Try
            _ribbonUIPassed = True
            Return ReadString("RibbonUI.xml")

        Catch ex As Exception

            _ribbonUIPassed = False
            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            Console.WriteLine(details)
            Return ""

        End Try

    End Function

#End Region

#Region "Ribbon Gui Trigger"

    Public Sub OnAction(ByVal control As IRibbonControl)
        Try

            Select Case control.Id
                Case "customButton1"
                    MessageBox.Show("This is the first sample button.")
                Case "customButton2"
                    MessageBox.Show("This is the second sample button.")
                Case Else
                    MessageBox.Show("Unkown Control Id: " + control.Id)

            End Select

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
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

        fileName = "ExcelAddinVB." + fileName

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
