Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports NetOffice
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<GuidAttribute("8A9ACFD2-77C2-4AE4-BBC4-9DD91D9C7E26"), ProgIdAttribute("AccessAddinVB4.TaskPaneAddin"), ComVisible(True)> _
Public Class Addin
    Implements IDTExtensibility2, Office.ICustomTaskPaneConsumer

    Private Shared ReadOnly _addinOfficeRegistryKey As String = "Software\\Microsoft\\Office\\Access\\AddIns\\"
    Private Shared ReadOnly _progId As String = "AccessAddinVB4.TaskPaneAddin"
    Private Shared ReadOnly _addinFriendlyName As String = "NetOffice Sample Addin in VB"
    Private Shared ReadOnly _addinDescription As String = "NetOffice Sample Addin with custom Task Pane"

    Shared _sampleControl As SampleControl
    Shared _accessApplication As Access.Application

    Public Shared ReadOnly Property Application() As Access.Application
        Get
            Return _accessApplication
        End Get
    End Property

#Region "ICustomTaskPaneConsumer Member"

    Public Sub CTPFactoryAvailable(ByVal CTPFactoryInst As Object) Implements NetOffice.OfficeApi.ICustomTaskPaneConsumer.CTPFactoryAvailable

        Try

            Dim ctpFactory As Office.ICTPFactory = New Office.ICTPFactory(_accessApplication, CTPFactoryInst)
            Dim taskPane As Office._CustomTaskPane = ctpFactory.CreateCTP(GetType(Addin).Assembly.GetName().Name + ".SampleControl", "NetOffice Sample Pane(VB4)", Type.Missing)
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft
            taskPane.Width = 300
            taskPane.Visible = True
            _sampleControl = taskPane.ContentControl
            ctpFactory.Dispose()

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#Region "IDTExtensibility2 Member"

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As Extensibility.ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        Try

            ' Initialize NetOffice
            NetOffice.Factory.Initialize()

            _accessApplication = New Access.Application(Nothing, Application)

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection

        Try

            If (Not IsNothing(_accessApplication)) Then
                _accessApplication.Dispose()
            End If

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown

    End Sub

#End Region

#Region "COM Register Functions"

    <ComRegisterFunctionAttribute()> _
    Public Shared Sub RegisterFunction(ByVal type As Type)
        Try
            ' add codebase value
            Dim thisAssembly As Assembly = Assembly.GetAssembly(GetType(Addin))
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

            ' add access addin key
            Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable")
            Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _progId)
            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _progId, True)
            rk.SetValue("LoadBehavior", CInt(3))
            rk.SetValue("FriendlyName", _addinFriendlyName)
            rk.SetValue("Description", _addinDescription)
            rk.Close()

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)
            Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _progId, False)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region

End Class
