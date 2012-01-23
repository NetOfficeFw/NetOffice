Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        '  NetOffice manages COM Proxies for you to avoid any kind of memory leaks
        '  and make sure your application instance removes from process list if you want.

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()
        ' now we have 2 new COM Proxies created.
        ' 
        ' the first proxy was created while accessing the Workbooks collection from application
        ' the second proxy was created by the Add() method from Workbooks and stored now in book
        ' with the application object we have 3 created proxies now. the workbooks proxy was created
        ' about application and the book proxy was created about the workbooks.
        ' NetOffice holds the proxies now in a list as follows:
        ' 
        ' Application
        '   + Workbooks
        '     + Workbook  
        ' 
        ' any object in NetOffice implements the IDisposible Interface.
        ' use the Dispose() Method to release an object. the method release all created child proxies too.

        application.Quit()
        application.Dispose()
        ' the application object is ouer root object
        ' dispose them release himself and any childs of application, in this case workbooks and workbook
        ' the excel instance are now removed from process list

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub linkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkDocEnglish.LinkClicked, linkFaqEnglish.LinkClicked, linkTeqFaqGerman.LinkClicked, linkTeqFaqEnglish.LinkClicked, linkTecDocGerman.LinkClicked, linkTecDocEnglish.LinkClicked, linkFaqGerman.LinkClicked, linkDocGerman.LinkClicked

        Dim ctrl As System.Windows.Forms.Control
        ctrl = sender
        System.Diagnostics.Process.Start(ctrl.Tag)

    End Sub

End Class
