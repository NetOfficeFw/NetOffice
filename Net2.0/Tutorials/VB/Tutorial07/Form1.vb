Imports System.Runtime.InteropServices

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        'Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook
        Dim book As Excel.Workbook = application.Workbooks.Add()

        Dim sheet As Excel.Worksheet = application.Workbooks(1).Worksheets(1)
        Dim sampleRange As Excel.Range = sheet.Cells(1, 1)

        'we set the COMVariant ColorIndex from Font of ouer sample range with the invoker class
        Invoker.PropertySet(sampleRange.Font, "ColorIndex", 1)

        ' creates a native unmanaged ComProxy with the invoker
        Dim comProxy As Object = Invoker.PropertyGet(application, "Workbooks")
        Marshal.ReleaseComObject(comProxy)

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
