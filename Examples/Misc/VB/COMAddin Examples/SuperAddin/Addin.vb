Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports NetOffice.Tools
Imports NetOffice.OfficeApi.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
'
' Multi-host addin example
'
<COMAddin("Super Addin Sample VB4", "Multi-host addin example", LoadBehavior.LoadAtStartup), Codebase>
<CustomUI("RibbonUI.xml", True), RegistryLocation(RegistrySaveLocation.CurrentUser)>
<ProgId("SuperAddinVB4.Connect"), Guid("B7561D9F-E3DE-49cd-B5FE-D812F8999EFD")>
<MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.PowerPoint, RegisterIn.Outlook, RegisterIn.Access, RegisterIn.MSProject)> _
Public Class Addin
    Inherits COMAddin

    Public Sub OnClickRibbonButton(ByVal control As Office.IRibbonControl)

        Try

            Select Case control.Id
                Case "customButton1"
                    MessageBox.Show(String.Format("Hosted in {0}", Application.InstanceFriendlyName))
                Case "customButton2"
                    MessageBox.Show(String.Format("Loading Time {0}", LoadingTimeElapsed))
            End Select

        Catch throwedException As Exception

            Utils.Dialog.ShowError(throwedException, "Unexpected state in SuperAddinVB4 OnClickRibbonButton")

        End Try

    End Sub

    Protected Overrides Sub OnError(ByVal methodKind As NetOffice.Tools.ErrorMethodKind, ByVal exception As System.Exception)

        Utils.Dialog.ShowError(exception, "Unexpected state in SuperAddinVB4 " + methodKind.ToString())

    End Sub

End Class
