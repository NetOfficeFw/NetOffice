Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook
        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Worksheets(1)

        Dim range As Excel.Range = sheet.Cells(1, 1)

        ' Style is defined as Variant in Excel Type Library and represents as object in NetOffice
        Dim style As Excel.Style = range.Style

        'variant types can be a scalar type, another way to us is 
        If (TypeName(range.Style) = "String") Then
            Dim myStyle As String = range.Style
        ElseIf (TypeName(range.Style) = "Style") Then
            Dim myStyle As Excel.Style = range.Style
        End If

        ' Name, Bold, Size are bool but defined as Variant and also converted to object
        style.Font.Name = "Arial"
        style.Font.Bold = True
        style.Font.Size = 14

        ' quit & dipose
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
