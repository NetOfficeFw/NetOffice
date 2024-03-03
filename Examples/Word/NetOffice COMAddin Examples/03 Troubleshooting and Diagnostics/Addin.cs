﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
using NetOffice.OfficeApi.Tools.Contribution;
/*
   Diagnostics Addin Example
*/
namespace Word03AddinCS4
{
    [COMAddin("Word03 Sample Addin CS4", "Diagnostics Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Word03AddinCS4.Connect"), Guid("FA094799-2FB7-4D61-B96F-CBCBFEA46C95"), Timestamp, ForceInitialize, Codebase]
    public class Addin : COMAddin
    {
        public Addin()
        {
            // Redirect console to System.Diagnostics.Trace and write a message
            Factory.Console.Mode = DebugConsoleMode.Trace;
            Factory.Console.WriteLine("Word03AddinCS4 has been started.");

            // Shared output want send all given console messages to a named pipe
            // ------------------------------------------------------------------
            //Factory.Console.EnableSharedOutput = false;
            //Factory.Console.Name = "Word03AddinCS4";

            OnStartupComplete += Addin_OnStartupComplete;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            // startup time elapsed
            Factory.Console.WriteLine("NetOffice has been initialized in {0}", Factory.InitializedTime);
            Factory.Console.WriteLine("Addin has been loaded completely in {0}", LoadingTimeElapsed);

            // Enable performance trace in Word to see all calls >= 3 milliseconds
            // See tutorials for further informations
            Factory.Settings.PerformanceTrace["NetOffice.WordApi"].IntervalMS = 3;
            Factory.Settings.PerformanceTrace["NetOffice.WordApi"].Enabled = true;
            Factory.Settings.PerformanceTrace.Alert += PerformanceTrace_Alert;

            // Setup a tray icon with context menu for available diagnostics
            Utils.Tray.Setup(true, "Addin Diagnostics", "Addin.ico");
            Utils.Tray.ShowBalloonTip(1000, "Addin Diagnostics", "Click here to see diagnostics", TrayToolTipIcon.Info);
            Utils.Tray.Menu.AutoClose = false;
            Utils.Tray.Menu.Items.Add<TrayMenuLabelItem>("Addin Diagnostics", true, "TrayMenuHeader.png");
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuMonitorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuItem>("Fetch documents");
            Utils.Tray.Menu.Items.Add<TrayMenuItem>("Dispose all application child proxies");
            Utils.Tray.Menu.Items.Add<TrayMenuSeparatorItem>();
            Utils.Tray.Menu.Items.Add<TrayMenuAutoCloseItem>("Enable Auto Close Menu");
            Utils.Tray.Menu.Items.Add<TrayMenuCloseItem>("Close Menu");
            Utils.Tray.Menu.ItemClick += Menu_ItemClick;

            // Check Word has been started from another program like: new Word.Application()
            bool automationMode = Utils.IsAutomation;

            // Check for admin permissions and Word is 2007 or higher in its version
            bool hasAdminPermissions = Utils.AdminPermissions;
            bool is2007OrHigher = Utils.ApplicationIs2007OrHigher;
        }

        private void Menu_ItemClick(object sender, TrayMenuItemsEventArgs args)
        {
            // See what happen in tray proxy live monitor
            if (args.Item.Text == "Fetch documents")
            {
                foreach (Word.Document doc in Application.Documents)
                {
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