Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        'Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook & attach close event trigger
        Dim book As Excel.Workbook = application.Workbooks.Add()

        Dim closeHandler As Excel.Workbook_BeforeCloseEventHandler = AddressOf Me.book_BeforeCloseEvent
        AddHandler book.BeforeCloseEvent, closeHandler

        ' we dispose the instance. the parameter false signals to api dont release the event listener
        ' set parameter to true and the event listener will stopped and you dont get events for the instance
        ' the DisposeChildInstances() method has the same method overload
        book.Close()
        book.Dispose(False)

        application.Quit()
        application.Dispose()
        ' 
        ' the application object is ouer root object
        ' dispose them release himself and any childs of application, in this case workbooks and workbook
        ' the excel instance are now removed from process list
        ' 

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Public Sub book_BeforeCloseEvent(ByRef Cancel As Boolean)

    End Sub

    Private Sub linkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkDocEnglish.LinkClicked, linkFaqEnglish.LinkClicked, linkTeqFaqGerman.LinkClicked, linkTeqFaqEnglish.LinkClicked, linkTecDocGerman.LinkClicked, linkTecDocEnglish.LinkClicked, linkFaqGerman.LinkClicked, linkDocGerman.LinkClicked

        Dim ctrl As System.Windows.Forms.Control
        ctrl = sender
        System.Diagnostics.Process.Start(ctrl.Tag)

    End Sub

End Class
