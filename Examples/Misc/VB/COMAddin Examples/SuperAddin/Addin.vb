Imports System.Reflection
Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Extensibility

Imports NetOffice
Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi
Imports Word = NetOffice.WordApi
Imports Outlook = NetOffice.OutlookApi
Imports PowerPoint = NetOffice.PowerPointApi
Imports Access = NetOffice.AccessApi


<Guid("C7E206D5-C681-4460-825E-6D44817BAD18"), ProgId("SuperAddinVB4.Addin"), ComVisible(True)> _
Public Class Addin
    Implements IDTExtensibility2, Office.IRibbonExtensibility

    Private Shared ReadOnly _progId As String = "SuperAddinVB4.Addin"
    Private Shared ReadOnly _addinFriendlyName As String = "NetOffice Sample Addin in VB"
    Private Shared ReadOnly _addinDescription As String = "NetOffice Sample Addin for multipe Office Applications"

    Private _application As COMObject
    Private _hostApplicationName As String

#Region "IDTExtensibility2 Members"

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As Extensibility.ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

        Try

            _application = Core.Default.CreateObjectFromComProxy(Nothing, Application)

            If (TypeName(_application) = "Excel.Application") Then
                _hostApplicationName = "Excel"
            ElseIf (TypeName(_application) = "Word.Application") Then
                _hostApplicationName = "Word"
            ElseIf (TypeName(_application) = "Outlook.Application") Then
                _hostApplicationName = "Outlook"
            ElseIf (TypeName(_application) = "PowerPoint.Application") Then
                _hostApplicationName = "PowerPoint"
            ElseIf (TypeName(_application) = "Access.Application") Then
                _hostApplicationName = "Access"
            End If

        Catch throwedException As Exception

            If (Not IsNothing(_hostApplicationName)) Then

                OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in OnConnection. ", throwedException)

            End If

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection

        Try

            If (Not IsNothing(_application)) Then
                _application.Dispose()
            End If

        Catch throwedException As Exception

            OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in OnDisconnection. ", throwedException)

        End Try

    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown

    End Sub

#End Region

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI

        Try

            Return ReadString("RibbonUI.xml")

        Catch throwedException As Exception

            OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in GetCustomUI. ", throwedException)
            Return ""

        End Try

    End Function

    Public Sub OnAction(ByVal control As Office.IRibbonControl)

        Try

            Dim messageString As String = String.Format("Thanks for click on a Ribbon." + vbNewLine + "HostApp is {0}.{1} Version:{2}", _
                                            TypeDescriptor.GetComponentName(_application.UnderlyingObject), _application.UnderlyingTypeName, Invoker.Default.PropertyGet(_application, "Version"))

            MessageBox.Show(messageString, _progId, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "OnAction " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

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

            Registry.ClassesRoot.CreateSubKey("CLSID\{" + type.GUID.ToString().ToUpper() + "}\Programmable")

            ' add bypass key
            ' http://support.microsoft.com/kb/948461
            key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}")
            Dim defaultValue As String = key.GetValue("")
            If (IsNothing(defaultValue)) Then
                key.SetValue("", "Office .NET Framework Lockback Bypass Key")
            End If
            key.Close()

            OfficeRegistry.CreateAddinKey("Excel", _progId, _addinFriendlyName, _addinDescription)
            OfficeRegistry.CreateAddinKey("Word", _progId, _addinFriendlyName, _addinDescription)
            OfficeRegistry.CreateAddinKey("Outlook", _progId, _addinFriendlyName, _addinDescription)
            OfficeRegistry.CreateAddinKey("PowerPoint", _progId, _addinFriendlyName, _addinDescription)
            OfficeRegistry.CreateAddinKey("Access", _progId, _addinFriendlyName, _addinDescription)

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)

            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Excel + _progId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Word + _progId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Outlook + _progId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.PowerPoint + _progId)
            OfficeRegistry.DeleteAddinKey(OfficeRegistry.Access + _progId)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

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

        Dim thisAssembly As Assembly = GetType(Addin).Assembly
        Dim ressourceStream As System.IO.Stream = thisAssembly.GetManifestResourceStream(thisAssembly.GetName().Name + "." + fileName)
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
