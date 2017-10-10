using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Tools;
/*
    Custom Property Page Addin Example
*/
namespace Outlook05AddinCS4
{
    [COMAddin("Outlook05 Sample Addin CS4", "Custom Property Page Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Outlook05AddinCS4.Connect"), Guid("E4BFB49B-633A-44B7-A568-68D82F836365"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            // Bring the option page to the Tools/Options menu
            Application.OptionsPagesAddEvent += Application_OptionsPagesAddEvent;

            /*
            // This is another way to bring a custom option page to
            // the Mail Folder TreeView on the left - context menu/properties
            // --------------------------------------------------------------
            Outlook.NameSpace mapi = Application.GetNamespace("MAPI") as Outlook.NameSpace;
            mapi.OptionsPagesAddEvent += Mapi_OptionsPagesAddEvent;
            */
        }

        private void Application_OptionsPagesAddEvent(Outlook.PropertyPages pages)
        {
            // we show the NetOffice Core settings in the option page
            pages.Add(new OptionPage(Factory), "Outlook05 Sample Addin CS4");
        }

        private void Mapi_OptionsPagesAddEvent(Outlook.PropertyPages pages, Outlook.MAPIFolder folder)
        {
            
        }
    }
}