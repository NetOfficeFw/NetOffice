using System;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi.Tools;

namespace Sample.Addin
{
    [COMAddin("Word Wikipedia Addin", "This Addin provides Wikipedia functionality", LoadBehavior.LoadAtStartup)]
    [ProgId("NOSample.Wikipedia"), Guid("56F843AD-ECB8-45D6-9E33-C0928BD2FB03")]
    [CustomPane(typeof(WikipediaPane), "Wikipedia - NetOffice Sample", true, PaneDockPosition.msoCTPDockPositionRight, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal, 520, 520)]
    public class ThisAddin : COMAddin
    {
        public ThisAddin()
        {
            this.OnConnection += new OnConnectionEventHandler(ThisAddin_OnConnection);
            this.OnStartupComplete += new OnStartupCompleteEventHandler(ThisAddin_OnStartupComplete);
        }

        private void ThisAddin_OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {

        }

        private void ThisAddin_OnStartupComplete(ref Array custom)
        {
            
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "Unexpected state in Wikipedia-Addin");
        }
    }
}
