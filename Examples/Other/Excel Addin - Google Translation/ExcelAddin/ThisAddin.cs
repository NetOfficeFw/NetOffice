using System;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;

namespace Sample.ExcelAddin
{
    /// <summary>
    /// The main addin for MS-Excel
    /// </summary>
    [COMAddin("Google Translation Addin", "This Addin provides Google Translation functionality", 3), Guid("fa65093e-8fd1-4e24-825a-11c00f1bcadf"), ProgId("NOSample.GoogleTranslation"), Tweak(true)]
    [CustomPane(typeof(TranslationPane), "NetOffice - Google Translation Sample", true, PaneDockPosition.msoCTPDockPositionBottom, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 150, 150)]
    public class ThisAddIn : COMAddin 
    {
        public ThisAddIn()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(ThisAddIn_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(ThisAddIn_OnDisconnection);
        }

        private void ThisAddIn_OnStartupComplete(ref Array custom)
        {
        }

        private void ThisAddIn_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "Unexpected state in Translation-Addin");
        }
    }
}
