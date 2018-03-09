Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial10
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' Enable and trigger trace alert
        NetOffice.Settings.Default.PerformanceTrace.Enabled = True
        Dim traceHandler As PerformanceTrace.PerformanceAlertEventHandler = AddressOf TraceAlert
        AddHandler NetOffice.Settings.Default.PerformanceTrace.Alert, traceHandler

        ' Criteria 1
        ' Enable performance trace in excel generaly. set interval limit to 100ms to see all actions there need >= 100 milliseconds
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi").Enabled = True
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi").IntervalMS = 100

        ' Criteria 2
        ' Enable additional performance trace for all members of WorkSheet in excel. set interval limit to 20ms to see all actions there need >=20 milliseconds
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi", "Worksheet").Enabled = True
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi", "Worksheet").IntervalMS = 20

        ' Criteria 3
        ' Enable additional performance trace for WorkSheet Range property in excel. set interval limit to 0ms to see all calls anywhere
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi", "Worksheet", "Range").Enabled = True
        NetOffice.Settings.Default.PerformanceTrace("NetOffice.ExcelApi", "Worksheet", "Range").IntervalMS = 0

        ' do some stuff
        Dim application As New Excel.Application()
        application.DisplayAlerts = False
        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Sheets.Add()
        For index = 1 To 5
            Dim range As Excel.Range = sheet.Range("A" + index.ToString())
            range.Value = "Test123"
            range(1, 1).Value = "Test234"
        Next

        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public Sub TraceAlert(sender As NetOffice.PerformanceTrace, args As NetOffice.PerformanceTrace.PerformanceAlertEventArgs)

        Console.WriteLine("{0} {1}:{2} in {3} Milliseconds ({4} Ticks)", args.CallType, args.EntityName, args.MethodName, args.TimeElapsedMS, args.Ticks)

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial10"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Measure Performance"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return FormMain.DocumentationBase & "Tutorial10_EN_VB.html"
        End Get
    End Property

End Class