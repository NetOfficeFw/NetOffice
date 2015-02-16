using System;
using System.Runtime.InteropServices;

using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;

namespace Sample.ExcelAddin
{
    /// <summary>
    /// The main addin for MS-Excel. The Addin use the base class COMAddin from NetOffice.ExcelApi.Tools.
    /// </summary>
    [COMAddin("Google Translation Addin", "This Addin provides Google Translation functionality", 3), Guid("fa65093e-8fd1-4e24-825a-11c00f1bcadf"), ProgId("NOSample.GoogleTranslation"), Tweak(true)]
    [CustomPane(typeof(TranslationPane), "NetOffice - Google Translation Sample", true, PaneDockPosition.msoCTPDockPositionBottom, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 150, 150)]
    public class ThisAddIn : COMAddin 
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ThisAddIn()
        {
            // register some typical addin events
            this.OnConnection += new OnConnectionEventHandler(ThisAddIn_OnConnection);
            this.OnDisconnection += new OnDisconnectionEventHandler(ThisAddIn_OnDisconnection);
        }

        /// <summary>
        /// Called from Excel. This is the first time the Application property(Excel.Application) from the addin is applicable
        /// </summary>
        /// <param name="Application"></param>
        /// <param name="ConnectMode"></param>
        /// <param name="AddInInst"></param>
        /// <param name="custom"></param>
        private void ThisAddIn_OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
             
        }

        /// <summary>
        /// Called from Excel. The Application property(Excel.Application) from the addin, and task panes if exists, was automaticly disposed after these call
        /// </summary>
        /// <param name="RemoveMode"></param>
        /// <param name="custom"></param>
        private  void ThisAddIn_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }

        // Catch errors from our event triggers and COMAddin base methods here         
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            
        }
    }
}
