Imports ExampleBase
Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Example08
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Delegate Sub UpdateEventTextDelegate(ByVal message As String)
    Dim _updateDelegate As UpdateEventTextDelegate

    Public Sub New()

        InitializeComponent()
        _updateDelegate = New UpdateEventTextDelegate(AddressOf UpdateTextbox)

    End Sub

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' its an example with an own visual control
        ' checkout buttonStartExample_Click

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example08", "Beispiel08")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Using Events", "Verwenden von Ereignissen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Me
        End Get
    End Property

#End Region

#Region "Trigger"

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False
        excelApplication.Visible = True

        ' we register some events. note: the event trigger was called from excel, means an other Thread
        ' you can get event notifys from various objects: Application or Workbook or Worksheet for example

        Dim newWorkbookHandler As Excel.Application_NewWorkbookEventHandler = AddressOf Me.excelApplication_NewWorkbook
        AddHandler excelApplication.NewWorkbookEvent, newWorkbookHandler

        Dim beforeCloseHandler As Excel.Application_WorkbookBeforeCloseEventHandler = AddressOf Me.excelApplication_WorkbookBeforeClose
        AddHandler excelApplication.WorkbookBeforeCloseEvent, beforeCloseHandler

        Dim workbookActivateHandler As Excel.Application_WorkbookActivateEventHandler = AddressOf Me.excelApplication_WorkbookActivate
        AddHandler excelApplication.WorkbookActivateEvent, workbookActivateHandler

        Dim workbookDeactivateHandler As Excel.Application_WorkbookDeactivateEventHandler = AddressOf Me.excelApplication_WorkbookDeactivate
        AddHandler excelApplication.WorkbookDeactivateEvent, workbookDeactivateHandler

        Dim sheetActivateHandler As Excel.Application_SheetActivateEventHandler = AddressOf Me.excelApplication_SheetActivateEvent
        AddHandler excelApplication.SheetActivateEvent, sheetActivateHandler

        Dim sheetDeactivateHandler As Excel.Application_SheetDeactivateEventHandler = AddressOf Me.excelApplication_SheetDeactivateEvent
        AddHandler excelApplication.SheetDeactivateEvent, sheetDeactivateHandler

        ' add a new workbook, add a sheet and close
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets.Add()
        workBook.Close()

        excelApplication.Quit()
        excelApplication.Dispose()

    End Sub

    Private Sub excelApplication_SheetDeactivateEvent(ByVal Sh As COMObject)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event SheetDeactivate called."})
        Sh.Dispose()

    End Sub

    Private Sub excelApplication_SheetActivateEvent(ByVal Sh As COMObject)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event SheetActivate called."})
        Sh.Dispose()

    End Sub

    Private Sub excelApplication_NewWorkbook(ByVal Wb As Excel.Workbook)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event NewWorkbook called."})
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, ByRef Cancel As Boolean)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event WorkbookBeforeClose called."})
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookActivate(ByVal Wb As Excel.Workbook)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event WorkbookActivate called."})
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookDeactivate(ByVal Wb As Excel.Workbook)

        textBoxEvents.BeginInvoke(_updateDelegate, New Object() {"Event WorkbookDeactivate called."})
        Wb.Dispose()

    End Sub

    Private Sub UpdateTextbox(ByVal message As String)

        textBoxEvents.AppendText(message & vbNewLine)

    End Sub

#End Region

End Class
