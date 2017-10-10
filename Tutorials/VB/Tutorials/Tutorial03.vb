Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Tutorial03
    Implements ITutorial

    Dim _hostApplication As IHost
    Dim _application As Excel.Application

    Public Sub New()

        InitializeComponent()

        CreateHandle()
        Dim changeHandler As Core.ProxyCountChangedHandler = AddressOf Me.Factory_ProxyCountChanged
        AddHandler Core.Default.ProxyCountChanged, changeHandler

    End Sub

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this tutorial shows you 3 ways in NetOffice to see how many com proxies
        ' was currently alive in your application
        '
        ' 1.) the property: Int NetOffice.Core.ProxyCount
        ' 2.) the event: NetOffice.Core.ProxyCountChanged()
        ' 3.) the events: NetOffice.Core ProxyAdded, ProxyRemoved, ProxyCleared
        '     used from NetOffice.Contribution.Controls.InstanceMonitor

        ' Note Sometimes you may wondering why an instance Is disposed.
        ' For troubleshooting you can trigger ICOMObject.OnDispose event And see strack trace

    End Sub

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication
        instanceMonitor1.Factory = NetOffice.Core.Default

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

        If Not IsNothing(_application) Then

            _application.Quit()
            _application.Dispose()
            _application = Nothing

        End If

        instanceMonitor1.Factory = Nothing

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial03"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Observable COM proxies"
        End Get
    End Property

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Me
        End Get
    End Property


    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return FormMain.DocumentationBase & "Tutorial03_EN_VB.html"
        End Get
    End Property


    Private Sub buttonExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonExcel.Click

        If (IsNothing(_application)) Then
            ' start application
            _application = New Excel.Application()
            _application.DisplayAlerts = False
            buttonExcel.Text = "Quit Excel"
            buttonWorkbook.Enabled = True
            buttonAddins.Enabled = True
            buttonDisposeChildInstances.Enabled = True
        Else
            ' quit application
            _application.Quit()
            _application.Dispose()
            _application = Nothing
            buttonExcel.Text = "Start Excel"
            buttonWorkbook.Enabled = False
            buttonAddins.Enabled = False
            buttonDisposeChildInstances.Enabled = False
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

    Private Sub buttonDisposeChildInstances_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonDisposeChildInstances.Click

        'dispose all child instances from application
        _application.DisposeChildInstances()

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

    ' its possible the event comes from a different thread, the method is an invoke helper to avoid a CrossThreadException
    Private Sub UpdateLabel()

        labelProxyCount.Text = labelProxyCount.Tag

    End Sub

End Class
