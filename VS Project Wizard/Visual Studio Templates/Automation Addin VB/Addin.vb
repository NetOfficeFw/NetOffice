Imports System.Reflection
Imports Microsoft.Win32
Imports Extensibility
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
$usingItems$
<GuidAttribute("$randomGuid$"), ProgIdAttribute("$safeprojectname$.$safeitemname$"), ComVisible(TRUE)> _
Public Class Addin
    Implements IDTExtensibility2$ribbonImplement$

$ApplicationField$

    Public Sub New()
    
    End Sub

#Region "IDTExtensibility2 Members"
    
    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete
$classicUICreateCall$   
    End Sub

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

 $ApplicationConstruction$

	' If the addin not connected during startup, we call OnStartupComplete at hand
    	If Not ConnectMode = ext_ConnectMode.ext_cm_Startup Then
        	OnStartupComplete(custom)
	End If
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection        
 		
	'If this is not because of host shutdown(removed by user for example) we call OnBeginShutdown at hand
    	If Not RemoveMode = ext_DisconnectMode.ext_dm_HostShutdown Then
		OnBeginShutdown(custom)
	End If

$ApplicationDestroy$
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown
$classicUIRemoveCall$
    End Sub

#End Region
$ribbonUIImplementMethod$$classicUICreateRemoveMethod$
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

$registerCode$
        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)
                
            ' unregister addin in office
$unregisterCode$	
	Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region
$helperCode$
End Class
