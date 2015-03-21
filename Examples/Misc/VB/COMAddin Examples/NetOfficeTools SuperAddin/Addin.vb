Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports NetOffice.Tools
Imports NetOffice.OfficeApi.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<COMAddin("NetOfficeTools Super Addin Sample", "This Addin shows you how i can create a NO Tools based Addin and support multiple office products", 3)> _
<RegistryLocation(RegistrySaveLocation.CurrentUser), CustomUI("RibbonUI.xml", True)> _
<Guid("B7561D9F-E3DE-49cd-B5FE-D812F8999EFD"), ProgId("NOToolsSuperAddinVB4.Addin"), ComVisible(True), Tweak(True)> _
<MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.PowerPoint, RegisterIn.Outlook, RegisterIn.Access, RegisterIn.MSProject)> _
Public Class Addin
    Inherits COMAddin

    Public Sub OnAction(ByVal control As Office.IRibbonControl)

        Try

            Select Case control.Id
                Case "customButton1"
                    Utils.Dialog.ShowMessageBox("This is the first sample button. " + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
                Case "customButton2"
                    Utils.Dialog.ShowMessageBox("This is the second sample button. " + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
                Case Else
                    Utils.Dialog.ShowMessageBox("Unkown Control Id: " + control.Id, "NetOfficeTools.SuperAddinVB4", DialogResult.None)
            End Select

        Catch throwedException As Exception

            Utils.Dialog.ShowError(throwedException, "Unexpected state in SuperAddinVB4 OnAction")

        End Try

    End Sub

    Protected Overrides Sub OnError(ByVal methodKind As NetOffice.Tools.ErrorMethodKind, ByVal exception As System.Exception)

        Utils.Dialog.ShowError(exception, "Unexpected state in SuperAddinVB4 " + methodKind.ToString())

    End Sub

End Class
