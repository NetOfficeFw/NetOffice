Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.ExcelApi.Tools
Imports NetOffice.OfficeApi.Tools.Contribution
Imports Excel = NetOffice.ExcelApi
'
'Diagnostics Addin Example
'
<COMAddin("Excel03 Sample Addin VB4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Excel03AddinVB4.Connect"), Guid("C46F9E2B-D428-4451-8609-7CB1B33CC7FE"), ForceInitialize, Timestamp, Codebase>
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        'Redirect console to System.Diagnostics.Trace and write a message
        Factory.Console.Mode = DebugConsoleMode.Trace
        Factory.Console.WriteLine("Excel03AddinCS4 has been started.")

        'Shared output want send all given console messages to a named pipe
        '------------------------------------------------------------------
        'Factory.Console.EnableSharedOutput = false
        'Factory.Console.Name = "Excel03AddinCS4"

    End Sub

    Private Sub Addin_OnStartupComplete(ByRef custom As Array) Handles Me.OnStartupComplete

        ' startup time elapsed
        Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime)
        Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed)

        ' Enable performance trace in Excel to see all calls >= 3 milliseconds
        ' See tutorials for further informations
        Factory.Settings.PerformanceTrace("NetOffice.ExcelApi").IntervalMS = 3
        Factory.Settings.PerformanceTrace("NetOffice.ExcelApi").Enabled = True

        ' Setup a tray icon with context menu for available diagnostics
        Utils.Tray.Setup(True, "Addin Diagnostics", "Addin.ico")
        Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info)
        Utils.Tray.Menu.AutoClose = False
        Utils.Tray.Menu.Items.Add(Of TrayMenuLabelItem)("Addin Diagnostics", True, "TrayMenuHeader.png")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuMonitorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Fetch books and sheets")
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Dispose all application child proxies")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuAutoCloseItem)("Enable Auto Close Menu")
        Utils.Tray.Menu.Items.Add(Of TrayMenuCloseItem)("Close Menu")
        Dim handler As TrayMenuItemClickEventHandler = AddressOf Me.Menu_ItemClick
        AddHandler Utils.Tray.Menu.ItemClick, handler

        ' Check Excel has been started from another program like: new Excel.Application()
        Dim automationMode As Boolean = Utils.IsAutomation

        ' Check for admin permissions and excel is 2007 or higher in its version
        Dim hasAdminPermissions As Boolean = Utils.AdminPermissions
        Dim is2007OrHigher As Boolean = Utils.ApplicationIs2007OrHigher

    End Sub

    '
    ' This method is called when COMAddin base is unable to complete an operation
    '
    Protected Overrides Sub OnError(methodKind As ErrorMethodKind, exception As Exception)

        Utils.Dialog.ShowErrorDefault(methodKind, exception)

    End Sub

    Public Sub Menu_ItemClick(sender As Object, args As TrayMenuItemsEventArgs)

        ' See what happen in tray proxy live monitor
        If args.Item.Text = "Fetch books and sheets" Then

            For Each book As Excel.Workbook In Application.Workbooks

                For Each sheet As Excel.Worksheet In book.Sheets

                Next sheet

            Next book

        ElseIf args.Item.Text = "Dispose all application child proxies" Then

            Application.DisposeChildInstances()

        End If

    End Sub

End Class