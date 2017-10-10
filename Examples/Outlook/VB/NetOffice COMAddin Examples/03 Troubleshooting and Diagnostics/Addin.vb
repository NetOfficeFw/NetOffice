Imports NetOffice
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.OutlookApi.Tools
Imports NetOffice.OfficeApi.Tools.Contribution
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
'
'Diagnostics Addin Example
'
<COMAddin("Outlook03 Sample Addin VB4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Outlook03AddinVB4.Connect"), Guid("E91A73EE-9A4C-4A82-AE0A-EC3745D6F542"), ForceInitialize, Timestamp, Codebase>
<RequireShutdownNotification>
Public Class Addin
    Inherits COMAddin

    Public Sub New()

        'Redirect console to System.Diagnostics.Trace and write a message
        Factory.Console.Mode = DebugConsoleMode.Trace
        Factory.Console.WriteLine("Outlook03AddinCS4 has been started.")

        'Shared output want send all given console messages to a named pipe
        '------------------------------------------------------------------
        'Factory.Console.EnableSharedOutput = false
        'Factory.Console.Name = "Outlook03AddinCS4"

    End Sub

    Private Sub Addin_OnStartupComplete(ByRef custom As Array) Handles Me.OnStartupComplete

        ' startup time elapsed
        Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime)
        Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed)

        ' Enable performance trace in Excel to see all calls >= 3 milliseconds
        ' See tutorials for further informations
        Factory.Settings.PerformanceTrace("NetOffice.OutlookApi").IntervalMS = 3
        Factory.Settings.PerformanceTrace("NetOffice.OutlookApi").Enabled = True

        ' Setup a tray icon with context menu for available diagnostics
        Utils.Tray.Setup(True, "Addin Diagnostics", "Addin.ico")
        Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info)
        Utils.Tray.Menu.AutoClose = False
        Utils.Tray.Menu.Items.Add(Of TrayMenuLabelItem)("Addin Diagnostics", True, "TrayMenuHeader.png")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuMonitorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Fetch inbox folder(first 20)")
        Utils.Tray.Menu.Items.Add(Of TrayMenuItem)("Dispose all application child proxies")
        Utils.Tray.Menu.Items.Add(Of TrayMenuSeparatorItem)()
        Utils.Tray.Menu.Items.Add(Of TrayMenuAutoCloseItem)("Enable Auto Close Menu")
        Utils.Tray.Menu.Items.Add(Of TrayMenuCloseItem)("Close Menu")
        Dim handler As TrayMenuItemClickEventHandler = AddressOf Me.Menu_ItemClick
        AddHandler Utils.Tray.Menu.ItemClick, handler

        ' Check Outlook has been started from another program like: new Outlook.Application()
        Dim automationMode As Boolean = Utils.IsAutomation

        ' Check for admin permissions and Outlook is 2007 or higher in its version
        Dim hasAdminPermissions As Boolean = Utils.AdminPermissions
        Dim is2007OrHigher As Boolean = Utils.ApplicationIs2007OrHigher

    End Sub

    '
    ' This method is called when COMAddin base is unable to complete an operation
    '
    Protected Overrides Sub OnError(methodKind As ErrorMethodKind, exception As Exception)

        Utils.Dialog.ShowErrorDefault(methodKind, exception)

    End Sub

    Private Sub Addin_OnBeginShutdown(ByRef custom As Array) Handles Me.OnBeginShutdown

        ' Outlook 2010 And higher do Not fire this event by default
        ' RequireShutdownNotification attribute(on top) bring it back into action
        Console.WriteLine("https://msdn.microsoft.com/library/office/ee720183.aspx")

    End Sub

    Public Sub Menu_ItemClick(sender As Object, args As TrayMenuItemsEventArgs)

        ' See what happen in tray proxy live monitor
        If args.Item.Text = "Fetch inbox folder(first 20)" Then

            Dim outlookNS As Outlook._NameSpace = Application.GetNamespace("MAPI")
            Dim inboxFolder As Outlook.MAPIFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
            For index = 1 To inboxFolder.Items.Count
                Dim item As Object = inboxFolder.Items(index)
                If index = 20 Then Exit For
            Next

        ElseIf args.Item.Text = "Dispose all application child proxies" Then

            Application.DisposeChildInstances()

        End If

    End Sub

End Class