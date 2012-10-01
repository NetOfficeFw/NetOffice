Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports Tests.Core

Public Class Test06
    Implements ITestPackage

    Dim _sheetDeactivateEvent As Boolean
    Dim _sheetActivateEvent As Boolean
    Dim _newWorkbookEvent As Boolean
    Dim _workbookBeforeCloseEvent As Boolean
    Dim _workbookActivateEvent As Boolean
    Dim _workbookDeactivateEvent As Boolean

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Using events."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test06"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Excel"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Excel.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False
            application.Visible = True
            application.Workbooks.Add()
              
            ' we register some events. note: the event trigger was called from excel, means an other Thread
            ' remove the Quit() call below and check out more events if you want
            ' you can get event notifys from various objects: Application or Workbook or Worksheet for example
            Dim newWorkbookHandler As Excel.Application_NewWorkbookEventHandler = AddressOf Me.excelApplication_NewWorkbook
            AddHandler application.NewWorkbookEvent, newWorkbookHandler

            Dim beforeCloseHandler As Excel.Application_WorkbookBeforeCloseEventHandler = AddressOf Me.excelApplication_WorkbookBeforeClose
            AddHandler application.WorkbookBeforeCloseEvent, beforeCloseHandler

            Dim workbookActivateHandler As Excel.Application_WorkbookActivateEventHandler = AddressOf Me.excelApplication_WorkbookActivate
            AddHandler application.WorkbookActivateEvent, workbookActivateHandler

            Dim workbookDeactivateHandler As Excel.Application_WorkbookDeactivateEventHandler = AddressOf Me.excelApplication_WorkbookDeactivate
            AddHandler application.WorkbookDeactivateEvent, workbookDeactivateHandler

            Dim sheetActivateHandler As Excel.Application_SheetActivateEventHandler = AddressOf Me.excelApplication_SheetActivateEvent
            AddHandler application.SheetActivateEvent, sheetActivateHandler

            Dim sheetDeactivateHandler As Excel.Application_SheetDeactivateEventHandler = AddressOf Me.excelApplication_SheetDeactivateEvent
            AddHandler application.SheetDeactivateEvent, sheetDeactivateHandler

            RemoveHandler application.NewWorkbookEvent, newWorkbookHandler
            RemoveHandler application.WorkbookBeforeCloseEvent, beforeCloseHandler
            RemoveHandler application.WorkbookActivateEvent, workbookActivateHandler
            RemoveHandler application.WorkbookDeactivateEvent, workbookDeactivateHandler
            RemoveHandler application.SheetActivateEvent, sheetActivateHandler
            RemoveHandler application.SheetDeactivateEvent, sheetDeactivateHandler

            ' add a new workbook add a sheet and close
            ' add a new workbook
            Dim workBook As Excel.Workbook = application.Workbooks.Add()
            Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)
            workBook.Close()

            RemoveHandler application.SheetDeactivateEvent, sheetDeactivateHandler

            If (_newWorkbookEvent And _workbookBeforeCloseEvent And _sheetActivateEvent And _sheetDeactivateEvent And _workbookActivateEvent And _workbookDeactivateEvent) Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Dim errorMessage As String = ""
                If (Not _newWorkbookEvent) Then errorMessage += "NewWorkbookEvent failed "
                If (Not _workbookBeforeCloseEvent) Then errorMessage += "WorkbookBeforeCloseEvent failed "
                If (Not _sheetActivateEvent) Then errorMessage += "WorkbookActivateEvent failed "
                If (Not _sheetDeactivateEvent) Then errorMessage += "WorkbookDeactivateEvent failed "
                If (Not _workbookActivateEvent) Then errorMessage += "SheetActivateEvent failed "
                If (Not _workbookDeactivateEvent) Then errorMessage += "SheetDeactivateEvent failed "
                Return New TestResult(True, DateTime.Now.Subtract(startTime), errorMessage, Nothing, "")
            End If

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function

    Private Sub excelApplication_SheetDeactivateEvent(ByVal Sh As COMObject)

        _sheetDeactivateEvent = True
        Sh.Dispose()

    End Sub

    Private Sub excelApplication_SheetActivateEvent(ByVal Sh As COMObject)

        _sheetActivateEvent = True
        Sh.Dispose()

    End Sub

    Private Sub excelApplication_NewWorkbook(ByVal Wb As Excel.Workbook)

        _newWorkbookEvent = True
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, ByRef Cancel As Boolean)

        _workbookBeforeCloseEvent = True
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookActivate(ByVal Wb As Excel.Workbook)

        _workbookActivateEvent = True
        Wb.Dispose()

    End Sub

    Private Sub excelApplication_WorkbookDeactivate(ByVal Wb As Excel.Workbook)

        _workbookDeactivateEvent = True
        Wb.Dispose()

    End Sub
End Class
