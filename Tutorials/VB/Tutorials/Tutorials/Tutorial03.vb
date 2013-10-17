Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Tutorial03
    Implements ITutorial

    Dim _hostApplication As IHost
    Dim _application As Excel.Application

    Public Sub New()

        InitializeComponent()

        Dim changeHandler As Core.ProxyCountChangedHandler = AddressOf Me.Factory_ProxyCountChanged
        AddHandler Core.Default.ProxyCountChanged, changeHandler

    End Sub

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this example shows you both ways in NetOffice to see how many com proxies
        ' was currently alive in your application
        '
        ' 1.) the static property: int NetOffice.Factory.ProxyCount
        ' 2.) the static event: NetOffice.Factory.ProxyCountChanged

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial03"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Observable COM Proxy Count", "Die Anzahl COM Proxies überwachen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub ChangeLanguage(ByVal lcid As Integer) Implements TutorialsBase.ITutorial.ChangeLanguage

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

        If Not IsNothing(_application) Then

            _application.Quit()
            _application.Dispose()

        End If

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Me
        End Get
    End Property


    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial03_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial03_DE_VB")
        End Get
    End Property

#End Region

#Region "UI Trigger"

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

        Dim proxyCount As Integer = NetOffice.Core.Default.ProxyCount

        Dim book As Excel.Workbook = _application.Workbooks.Add()
        book.Worksheets.Add()

        Dim proxyCountAfterCreate As Integer = NetOffice.Core.Default.ProxyCount

        'dispose all child instances from application
        _application.DisposeChildInstances()

        Dim proxyCountAfterDispose As Integer = NetOffice.Core.Default.ProxyCount

        Dim message As String = String.Format(
                                       "ProxyCount before create is {0}{3}" +
                                       "ProxyCount after create is {1}{3}" +
                                       "ProxyCount after dispose all childs from application is {2}", proxyCount, proxyCountAfterCreate, proxyCountAfterDispose, Environment.NewLine)

        MessageBox.Show(message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

#End Region

    Private Sub Factory_ProxyCountChanged(ByVal proxyCount As Integer)

        If (labelProxyCount.InvokeRequired) Then
            labelProxyCount.Tag = proxyCount.ToString()
            Dim updateHandler As MethodInvoker = AddressOf Me.UpdateLabel
            labelProxyCount.Invoke(updateHandler)
        Else
            labelProxyCount.Text = proxyCount.ToString()
        End If

    End Sub

    ' its possible the event comes from a different thread, the method is an invoke helper to avoid a CrossThreadException
    Private Sub UpdateLabel()

        labelProxyCount.Text = labelProxyCount.Tag

    End Sub

End Class
