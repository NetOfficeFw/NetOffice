Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.PowerPointApi.Tools
Imports NetOffice.OfficeApi.Tools.Contribution
Imports PowerPoint = NetOffice.PowerPointApi
'
'Diagnostics Addin Example
'
<COMAddin("PowerPoint03 Sample Addin VB4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("PowerPoint03AddinVB4.Connect"), Guid("BD789035-8CAF-491B-9B86-26A951620643"), ForceInitialize, Timestamp, Codebase>
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        'Redirect console to System.Diagnostics.Trace and write a message
        Factory.Console.Mode = DebugConsoleMode.Trace
        Factory.Console.WriteLine("PowerPoint03AddinCS4 has been started.")

        'Shared output want send all given console messages to a named pipe
        '------------------------------------------------------------------
        'Factory.Console.EnableSharedOutput = false
        'Factory.Console.Name = "PowerPoint03AddinCS4"

    End Sub

    Private Sub Addin_OnStartupComplete(ByRef custom As Array) Handles Me.OnStartupComplete

        ' startup time elapsed
        Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime)
        Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed)

        ' Enable performance trace in PowerPoint to see all calls >= 3 milliseconds
        ' See tutorials for further informations
        Factory.Settings.PerformanceTrace("NetOffice.PowerPointApi").IntervalMS = 3
        Factory.Settings.PerformanceTrace("NetOffice.PowerPointApi").Enabled = True

        ' Setup a tray icon with context menu for available diagnostics
        Utils.Tray.Setup(True, "Addin Diagnostics", "Addin.ico")
        Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info)
        Utils.Tray.Menu.AutoClose = False
        Utils.Tray.Menu.Items.Add(Of TrayMenuLabelItem)("Addin Diagnostics", True, "TrayMenuHeader.png")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuMonitorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Fetch Presentation")
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Dispose all application child proxies")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuAutoCloseItem)("Enable Auto Close Menu")
        Utils.Tray.Menu.Items.Add(Of TrayMenuCloseItem)("Close Menu")
        Dim handler As TrayMenuItemClickEventHandler = AddressOf Me.Menu_ItemClick
        AddHandler Utils.Tray.Menu.ItemClick, handler

        ' Check PowerPoint has been started from another program like: new Excel.PowerPoint()
        Dim automationMode As Boolean = Utils.IsAutomation

        ' Check for admin permissions and PowerPoint is 2007 or higher in its version
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
        If args.Item.Text = "Fetch Presentation" Then

            For Each pres As PowerPoint.Presentation In Application.Presentations

                For Each slide As PowerPoint.Slide In pres.Slides

                Next slide

            Next pres

        ElseIf args.Item.Text = "Dispose all application child proxies" Then

            Application.DisposeChildInstances()

        End If

    End Sub

End Class