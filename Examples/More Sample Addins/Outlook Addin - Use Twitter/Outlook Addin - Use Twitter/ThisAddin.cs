using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.OutlookApi.Tools;

namespace Sample.Addin
{
    /// <summary>
    /// The main addin for MS-Word. The Addin use the base class COMAddin from NetOffice.OutlookApi.Tools.
    /// </summary>
    [ProgId("NOSample.Twitter"), Guid("60875BD5-C5A5-4315-8954-BEEF3112DA82")]
    [COMAddin("Outlook Twitter Addin", "This Addin provides Twitter functionality", 3), Tweak(true)]
    [CustomPane(typeof(TwitterPane), "Twitter (powered by Linq2Twitter)", true, PaneDockPosition.msoCTPDockPositionRight, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal, 300, 300)]
    public class ThisAddin : COMAddin
    {
        public ThisAddin()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(ThisAddin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(ThisAddin_OnDisconnection);
        }

        private void ThisAddin_OnStartupComplete(ref Array custom)
        {

        }

        private void ThisAddin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "Unexpected state in Twitter-Addin");
        }
    }
}
