Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' Initialize NetOffice
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()

        ' ActiveSheet is defined as unkown Proxy in Excel Type Library, it can have multiple times at runtime
        ' but its always a COM Proxy, never a scalar type like bool or int. 
        ' In VBA oder PIA its converted to object, in NetOffice its represents as COMObject
        ' All NetOffice classes inherited COMObject
        Dim sheet As COMObject = application.ActiveSheet
        If (TypeName(sheet) = "Worksheet") Then
            Dim activeSheet As Excel.Worksheet = sheet
        End If

        '3 basic properties of COMObject
        Dim proxy As Object = sheet.UnderlyingObject
        Dim proxyClassName As String = sheet.UnderlyingTypeName
        Dim isDisposed As Boolean = sheet.IsDisposed

        application.Quit()
        application.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub linkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkDocEnglish.LinkClicked, linkFaqEnglish.LinkClicked, linkTeqFaqGerman.LinkClicked, linkTeqFaqEnglish.LinkClicked, linkTecDocGerman.LinkClicked, linkTecDocEnglish.LinkClicked, linkFaqGerman.LinkClicked, linkDocGerman.LinkClicked

        Dim ctrl As System.Windows.Forms.Control
        ctrl = sender
        System.Diagnostics.Process.Start(ctrl.Tag)

    End Sub

End Class
