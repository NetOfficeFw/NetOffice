Imports System.Reflection
Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports NetOffice
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<GuidAttribute("CEED1136-9063-4621-8BEB-A4DA8B35D89B"), ProgIdAttribute("WordAddinVB4.RibbonAddin"), ComVisible(True)> _
Public Class Addin
    Implements IDTExtensibility2, Office.IRibbonExtensibility

    Private Shared ReadOnly _addinOfficeRegistryKey As String = "Software\\Microsoft\\Office\\Word\\AddIns\\"
    Private Shared ReadOnly _prodId As String = "WordAddinVB4.RibbonAddin"
    Private Shared ReadOnly _addinFriendlyName As String = "NetOffice Sample Addin in VB"
    Private Shared ReadOnly _addinDescription As String = "NetOffice Sample Addin with custom Ribbon UI"

    Dim _wordApplication As Word.Application = Nothing

#Region "IDTExtensibility2 Members"

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        Try

            _wordApplication = New Word.Application(Nothing, Application)

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection
        Try

            If (Not IsNothing(_wordApplication)) Then
                _wordApplication.Dispose()
            End If

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

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

            ' add word addin key
            Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable")
            Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _prodId)
            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _prodId, True)
            rk.SetValue("LoadBehavior", CInt(3))
            rk.SetValue("FriendlyName", _addinFriendlyName)
            rk.SetValue("Description", _addinDescription)
            rk.Close()

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _addinFriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    <ComUnregisterFunctionAttribute()> _
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)
            Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _prodId, False)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ByVal RibbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI

        Try

            Return ReadString("RibbonUI.xml")

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured in GetCustomUI." + details, "GetCustomUI " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""

        End Try

    End Function

#End Region

#Region "Ribbon Gui Trigger"

    Public Sub OnAction(ByVal control As Office.IRibbonControl)
        Try

            Select Case control.Id
                Case "customButton1"
                    MessageBox.Show("This is the first sample button.", _prodId)
                Case "customButton2"
                    MessageBox.Show("This is the second sample button.", _prodId)
                Case Else
                    MessageBox.Show("Unkown Control Id: " + control.Id, _prodId)

            End Select

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "OnAction " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error)

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
