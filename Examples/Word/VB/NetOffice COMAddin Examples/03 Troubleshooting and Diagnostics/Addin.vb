Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
Imports NetOffice.OfficeApi.Tools.Contribution
Imports Word = NetOffice.WordApi
'
'Diagnostics Addin Example
'
<COMAddin("Word03 Sample Addin VB4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Word03AddinVB4.Connect"), Guid("E4E5EBFD-24B3-47DB-B572-22C977266330"), ForceInitialize, Timestamp, Codebase>
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        'Redirect console to System.Diagnostics.Trace and write a message
        Factory.Console.Mode = DebugConsoleMode.Trace
        Factory.Console.WriteLine("Word03AddinCS4 has been started.")

        'Shared output want send all given console messages to a named pipe
        '------------------------------------------------------------------
        'Factory.Console.EnableSharedOutput = false
        'Factory.Console.Name = "Word03AddinCS4"

    End Sub

    Private Sub Addin_OnStartupComplete(ByRef custom As Array) Handles Me.OnStartupComplete

        ' startup time elapsed
        Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime)
        Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed)

        ' Enable performance trace in Excel to see all calls >= 3 milliseconds
        ' See tutorials for further informations
        Factory.Settings.PerformanceTrace("NetOffice.WordApi").IntervalMS = 3
        Factory.Settings.PerformanceTrace("NetOffice.WordApi").Enabled = True

        ' Setup a tray icon with context menu for available diagnostics
        Utils.Tray.Setup(True, "Addin Diagnostics", "Addin.ico")
        Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info)
        Utils.Tray.Menu.AutoClose = False
        Utils.Tray.Menu.Items.Add(Of TrayMenuLabelItem)("Addin Diagnostics", True, "TrayMenuHeader.png")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuMonitorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Fetch documents")
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Dispose all application child proxies")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuAutoCloseItem)("Enable Auto Close Menu")
        Utils.Tray.Menu.Items.Add(Of TrayMenuCloseItem)("Close Menu")
        Dim handler As TrayMenuItemClickEventHandler = AddressOf Me.Menu_ItemClick
        AddHandler Utils.Tray.Menu.ItemClick, handler

        ' Check Word has been started from another program like: new Word.Application()
        Dim automationMode As Boolean = Utils.IsAutomation

        ' Check for admin permissions and Word is 2007 or higher in its version
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
        If args.Item.Text = "Fetch documents" Then

            For Each book As Word.Document In Application.Documents

            Next book

        ElseIf args.Item.Text = "Dispose all application child proxies" Then

            Application.DisposeChildInstances()

        End If

    End Sub

End Class