using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi.Tools;
using NetOffice.OfficeApi.Tools.Contribution;
/*
   Diagnostics Addin Example
*/
namespace Outlook03AddinCS4
{
    [COMAddin("Outlook03 Sample Addin CS4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Outlook03AddinCS4.Connect"), Guid("A39338C1-55FC-41CA-B69A-772A805231E9"), Timestamp, ForceInitialize, Codebase]
    [RequireShutdownNotification]
    public class Addin : COMAddin
    {
        public Addin()
        {           
            // Redirect console to System.Diagnostics.Trace and write a message
            Factory.Console.Mode = DebugConsoleMode.Trace;
            Factory.Console.WriteLine("Outlook03AddinCS4 has been started.");

            // Shared output want send all given console messages to a named pipe
            // ------------------------------------------------------------------
            //Factory.Console.EnableSharedOutput = false;
            //Factory.Console.Name = "Outlook03AddinCS4";

            OnStartupComplete += Addin_OnStartupComplete;
            OnBeginShutdown += Addin_OnBeginShutdown;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            // startup time elapsed
            Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime);
            Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed);

            // Enable performance trace in Outlook to see all calls >= 3 milliseconds
            // See tutorials for further informations
            Factory.Settings.PerformanceTrace["NetOffice.OutlookApi"].IntervalMS = 3;
            Factory.Settings.PerformanceTrace["NetOffice.OutlookApi"].Enabled = true;
            Factory.Settings.PerformanceTrace.Alert += PerformanceTrace_Alert;

            // Setup a tray icon with context menu for available diagnostics
            Utils.Tray.Setup(true, "Addin Diagnostics", "Addin.ico");
            Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info);
            Utils.Tray.Menu.AutoClose = false;
            Utils.Tray.Menu.Items.Add<TrayMenuLabelItem>("Addin Diagnostics", true, "TrayMenuHeader.png");
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuMonitorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuItem>("Fetch inbox(first 20)");
            Utils.Tray.Menu.Items.Add<TrayMenuItem>("Dispose all application child proxies");
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuAutoCloseItem>("Enable Auto Close Menu");
            Utils.Tray.Menu.Items.Add<TrayMenuCloseItem>("Close Menu");
            Utils.Tray.Menu.ItemClick += Menu_ItemClick;

            // Check Outlook has been started from another program like: new Outlook.Application()
            bool automationMode = Utils.IsAutomation;

            // Check for admin permissions and Outlook is 2007 or higher in its version
            bool hasAdminPermissions = Utils.AdminPermissions;
            bool is2007OrHigher = Utils.ApplicationIs2007OrHigher;
        }

        private void Addin_OnBeginShutdown(ref Array custom)
        {
            // Outlook 2010 and higher do not fire this event by default
            // RequireShutdownNotification attribute(on top) bring it back into action
            Console.WriteLine("https://msdn.microsoft.com/library/office/ee720183.aspx");
        }

        private void Menu_ItemClick(object sender, TrayMenuItemsEventArgs args)
        {
            // See what happen in tray proxy live monitor
            if (args.Item.Text == "Fetch inbox(first 20)")
            {
                Outlook._NameSpace outlookNS = Application.GetNamespace("MAPI");
                Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                for (int i = 1; i < inboxFolder.Items.Count; i++)
                {
                    object item = inboxFolder.Items[i];
                    if (i == 20)
                        break;
                }
            }
            else if (args.Item.Text == "Dispose all application child proxies")
            {
                Application.DisposeChildInstances();
            }
        }

        /*
            This method is called when COMAddin base is unable to complete an operation
        */
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowErrorDefault(methodKind, exception);
        }

        private void PerformanceTrace_Alert(PerformanceTrace sender, PerformanceTrace.PerformanceAlertEventArgs args)
        {            
            Factory.Console.WriteLine("PerformanceTrace Alert: {0}", args);
        }
    }
}