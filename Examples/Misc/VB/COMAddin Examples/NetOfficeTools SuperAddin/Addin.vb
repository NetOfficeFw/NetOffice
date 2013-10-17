Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports NetOffice.Tools
Imports NetOffice.OfficeApi.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

<COMAddin("NetOfficeTools Super Addin Sample", "This Addin shows you how i can create a NO Tools based Addin and support multiple office products", 3)> _
<RegistryLocation(RegistrySaveLocation.CurrentUser), CustomUI("NetOfficeTools.SuperAddinVB4.RibbonUI.xml")> _
<Guid("B7561D9F-E3DE-49cd-B5FE-D812F8999EFD"), ProgId("NOToolsSuperAddinVB4.Addin"), ComVisible(True), Tweak(True)> _
<MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.PowerPoint, RegisterIn.Outlook, RegisterIn.Access, RegisterIn.MSProject)> _
Public Class Addin
    Inherits COMAddin

    Public Sub OnAction(ByVal control As Office.IRibbonControl)

        Try

            Select Case control.Id
                Case "customButton1"
                    MessageBox.Show("This is the first sample button. " & Application.FriendlyTypeName, "NetOfficeTools.SuperAddinVB4")
                Case "customButton2"
                    MessageBox.Show("This is the second sample button. " & Application.FriendlyTypeName, "NetOfficeTools.SuperAddinVB4")
                Case Else
                    MessageBox.Show("Unkown Control Id: " + control.Id, "NetOfficeTools.SuperAddinVB4")

            End Select

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured in OnAction." + details, "NetOfficeTools.SuperAddinVB4", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class
