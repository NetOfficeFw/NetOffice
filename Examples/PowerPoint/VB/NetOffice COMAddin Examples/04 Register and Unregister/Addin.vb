Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Tools.Contribution
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Tools

'Register Addin Example
'
<COMAddin("PowerPoint04 Sample Addin VB4", "Register Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("PowerPoint04AddinVB4.Connect"), Guid("EB92B724-E88E-4614-9BB6-50F360E6760B"), Codebase, Timestamp>
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

        DialogUtils.ShowRegisterError("PowerPoint04AddinVB4", methodKind, exception)

    End Sub

End Class