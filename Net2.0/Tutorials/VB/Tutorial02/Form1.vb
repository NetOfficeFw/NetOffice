Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize Api COMObject Support 
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Worksheets.Add()
        ' we have 5 created proxies now in proxy table as follows
        ' 
        ' Application
        '  + Workbooks
        '     + Workbook  
        '        + Worksheets  
        '           + Worksheet  
        '

        ' we dispose the child instances of book
        book.DisposeChildInstances()

        ' we have 3 created proxies now, the childs from book are disposed
        ' 
        ' Application
        '   + Workbooks
        '    + Workbook  
        '

        application.Quit()
        application.Dispose()
        '
        'the Dispose() call for application release the instance and created childs Workbooks and Workbook
        '

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub linkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkDocEnglish.LinkClicked, linkFaqEnglish.LinkClicked, linkTeqFaqGerman.LinkClicked, linkTeqFaqEnglish.LinkClicked, linkTecDocGerman.LinkClicked, linkTecDocEnglish.LinkClicked, linkFaqGerman.LinkClicked, linkDocGerman.LinkClicked

        Dim ctrl As System.Windows.Forms.Control
        ctrl = sender
        System.Diagnostics.Process.Start(ctrl.Tag)

    End Sub

End Class
