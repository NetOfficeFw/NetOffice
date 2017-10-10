Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Tools.Contribution
Imports Access = NetOffice.AccessApi
Imports NetOffice.AccessApi.Tools

'Register Addin Example
'
<COMAddin("Access04 Sample Addin VB4", "Register Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Access04AddinVB4.Connect"), Guid("FAF18C76-9BF1-4763-BFCB-4AEB7C5EA893"), Codebase, Timestamp>
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

        DialogUtils.ShowRegisterError("Access04AddinVB4", methodKind, exception)

    End Sub

End Class