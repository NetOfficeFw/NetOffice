Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Imports NetOffice
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

'/*
' *    This project shows you the Tweak attribute in NetOffice.
' *    You can set the Tweak attribute to set/manipulate NetOffice options or your own options at runtime.
' *    This can be very helpful for developers, may troubleshooting or diagnostics, whatever.
' *    All Tweak settings has to be stored as string value in the current office addin registry key. For example:(HKEY_CurrentUser\Sofware\Microsoft\Office\%Application%\Addins\YourAddin)
' *    You find all possible NetOffice default tweak settings here: http://netoffice.codeplex.com/wikipage?=Tweaks"
' *    In this project you learn how you get control about tweaks and implement your own tweaks in an easy way.*    
' */

<COMAddin("NetOfficeVB4 Sample Word Addin", "This Addin shows you the COMAddin tweak option from the NetOffice Tools", 3)> _
<Guid("CC345443-D312-4B11-BB38-B3EAA0F776B5"), ProgId("TweakWordVB4.Addin"), Tweak(True)> _
Public Class Addin
    Inherits COMAddin

    ' This method was called for all (currently found) tweaks while startup. This means the NetOffice tweaks and your own tweaks.
    ' You have to decide the tweak is allowed or not. Please keep in your mind: All NetOffice tweak names starts with 'NO'
    Protected Overrides Function AllowApplyTweak(ByVal name As String, ByVal value As String) As Boolean

        ' we accept all tweaks
        Return True

    End Function

    ' This method was called from IExtensibility2.OnStartupComplete for all your custom tweaks if its allowed(see AllowApplyTweak)
    Protected Overrides Sub ApplyCustomTweak(ByVal name As String, ByVal value As String)

        If (name = "ShowTray" And value = "yes") Then
            Utils.Tray.Text = "TweakWordVB4.Addin"
            Utils.Tray.Visible = True
        End If

    End Sub

    ' This method was called from IExtensibility2.OnDisconnection for all your allowed custom aplied tweaks to remove or unload them.
    ' Please keep in your mind: the method is never called in state of unexpected termination. you have no warranties for the method.
    Protected Overrides Sub DisposeCustomTweak(ByVal name As String, ByVal value As String)

    End Sub

    ' We set some default- and custom tweaks in the register method.
    ' Please note: Installers like .msi or other doesnt call the static register methods for your (managed) addin while un-/registration.
    ' You have to set these entries at hand in the corresponding deployment project.
    <RegisterFunction(RegisterMode.CallAfter)> _
    Public Shared Sub Register(ByVal type As Type, ByVal registerCall As RegisterCall)

        ' SetTweakPersistenceEntry sets the key for you in the current registry key.
        ' We set a custom tweak and a Netoffice default tweak.
        SetTweakPersistenceEntry(type, "ShowTray", "yes", False)
        SetTweakPersistenceEntry(type, "NOConsoleMode", "trace", False)

    End Sub

End Class
