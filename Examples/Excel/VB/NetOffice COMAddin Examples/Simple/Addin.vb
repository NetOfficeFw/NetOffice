Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.ExcelApi.Tools
Imports NetOffice.ExcelApi

'/*
'  * This project shows you the COMAddin base class from the NetOffice tools.
'  * Its designed to reduce infrastructure code from your own.
'  * You have to set some attributes and thats all. 
'  * You see also the host application is available as class instance property. no need for dispose here because the base class do this for you while shutdown.
'*/

<COMAddin("NetOfficeVB4 Sample Excel Addin", "This Addin shows you the COMAddin base class from the NetOffice Tools", 3)> _
<Guid("B5CBBECD-4DEE-4A61-AD34-E9B8618452B0"), ProgId("SimpleExcelVB4.Addin")> _
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' show the host application version
        Dim hostVersion As String = String.Format("Host Application Version is:{0}", Me.Application.Version)
        Utils.Dialog.ShowMessageBox(hostVersion, MessageBoxIcon.Information, DialogResult.OK)

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class
