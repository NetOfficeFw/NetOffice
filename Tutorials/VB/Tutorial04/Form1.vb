Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Dim _application As Excel.Application

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Initialize Api COMObject Support 
        LateBindingApi.Core.Factory.Initialize()

        Dim changeHandler As Factory.ProxyCountChangedHandler = AddressOf Me.Factory_ProxyCountChanged
        AddHandler Factory.ProxyCountChanged, changeHandler

    End Sub


    Private Sub Factory_ProxyCountChanged(ByVal proxyCount As Integer)

        If (labelProxyCount.InvokeRequired) Then
            labelProxyCount.Tag = proxyCount.ToString()
            Dim updateHandler As MethodInvoker = AddressOf Me.UpdateLabel
            labelProxyCount.Invoke(updateHandler)
        Else
            labelProxyCount.Text = proxyCount.ToString()
        End If

    End Sub

    Private Sub buttonExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonExcel.Click

        If (IsNothing(_application)) Then
            ' start application
            _application = New Excel.Application()
            _application.DisplayAlerts = False
            buttonExcel.Text = "Quit Excel"
            buttonWorkbook.Enabled = True
            buttonAddins.Enabled = True
            buttonAddRemoveWorkbook.Enabled = True
        Else
            ' quit application
            _application.Quit()
            _application.Dispose()
            _application = Nothing
            buttonExcel.Text = "Start Excel"
            buttonWorkbook.Enabled = False
            buttonAddins.Enabled = False
            buttonAddRemoveWorkbook.Enabled = False
        End If

    End Sub

    Private Sub buttonWorkbook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonWorkbook.Click

        ' 2 new proxies, the workbooks proxy(implicit) and the new workbook from Add()
        If (Not IsNothing(_application)) Then
            _application.Workbooks.Add()
        End If

    End Sub

    Private Sub buttonAddins_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonAddins.Click

        If (Not IsNothing(_application)) Then
            '1 new enumerator proxy and 1 new proxy for any Addin
            Dim item As Excel.AddIn
            For Each item In _application.AddIns
                Console.WriteLine(item.Name)
            Next item
        End If

    End Sub

    Private Sub buttonAddRemoveWorkbook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonAddRemoveWorkbook.Click

        ' add a new worbook and a new worksheet to the workbook
        ' the worksheet is a child proxy from worbook, after dispose the workbook
        ' creates 4 new proxies
        ' the open proxy count is the same as before

        Dim proxyCount As Integer = LateBindingApi.Core.Factory.ProxyCount

        Dim book As Excel.Workbook = _application.Workbooks.Add()
        book.Worksheets.Add()

        Dim proxyCountAfterCreate As Integer = LateBindingApi.Core.Factory.ProxyCount

        'dispose all child instances from application
        _application.DisposeChildInstances()

        Dim proxyCountAfterDispose As Integer = LateBindingApi.Core.Factory.ProxyCount

        Dim message As String = String.Format("Method creates a new Workbook with 1 new Worksheet" & vbNewLine & _
                                                "ProxyCount before create is {0}" & vbNewLine & _
                                                "ProxyCount after create is {1}" & vbNewLine & _
                                                "ProxyCount after dispose Workbook is {2}", proxyCount, proxyCountAfterCreate, proxyCountAfterDispose)

        MessageBox.Show(message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub UpdateLabel()

        labelProxyCount.Text = labelProxyCount.Tag

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        If (Not IsNothing(_application)) Then
            _application.Quit()
            _application.Dispose()
            _application = Nothing
        End If

    End Sub

    Private Sub linkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkDocEnglish.LinkClicked, linkFaqEnglish.LinkClicked, linkTeqFaqGerman.LinkClicked, linkTeqFaqEnglish.LinkClicked, linkTecDocGerman.LinkClicked, linkTecDocEnglish.LinkClicked, linkFaqGerman.LinkClicked, linkDocGerman.LinkClicked

        Dim ctrl As System.Windows.Forms.Control
        ctrl = sender
        System.Diagnostics.Process.Start(ctrl.Tag)

    End Sub
End Class
