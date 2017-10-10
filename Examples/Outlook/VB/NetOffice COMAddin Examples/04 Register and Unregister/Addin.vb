Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Tools.Contribution
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Tools

'Register Addin Example
'
<COMAddin("Outlook04 Sample Addin VB4", "Register Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Outlook04AddinVB4.Connect"), Guid("A3C84E99-378C-4018-BA95-C20F013ECE10"), Codebase, Timestamp>
<RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser)>
Public Class Addin
    Inherits COMAddin

    <UnRegisterFunction(RegisterMode.CallAfter)> ' We want that NetOffice call this method after register
    Private Shared Sub Register(type As Type, registerCall As RegisterCall, scope As InstallScope, keyState As OfficeRegisterKeyState)

    End Sub

    <RegisterFunction(RegisterMode.CallBeforeAndAfter)> ' We want that NetOffice call this method before and after unregister
    Private Shared Sub Unregister(type As Type, registerCall As RegisterCall, scope As InstallScope, keyState As OfficeRegisterKeyState)

    End Sub

    ' An unexpected error occured in register or unregister action
    <RegisterErrorHandler>
    Private Shared Sub RegisterError(methodKind As RegisterErrorMethodKind, exception As Exception)

        DialogUtils.ShowRegisterError("Outlook04AddinVB4", methodKind, exception)

    End Sub

End Class