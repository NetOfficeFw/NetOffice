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
    /// Learn more about the NetOffice Tools namespace: http://netoffice.codeplex.com/wikipage?title=Tools_EN
    /// </summary>
    [Guid("fa65093e-8fd1-4e24-825a-11c00f1bcadf"), ProgId("Sample.ExcelAddin"), Tweak(true)]
    [COMAddin("Google Translation Addin", "This Addin provides Google Translation functionality", 3)]
    public class ThisAddIn : COMAddin 
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ThisAddIn()
        {
            // we create the taskpane
            TaskPanes.Add(typeof(TranslationPane), "Google Translation - NetOffice Local Shared Data Sample");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical;
            TaskPanes[0].Height = 150;
            TaskPanes[0].Visible = true;

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
        void ThisAddIn_OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
             
        }

        /// <summary>
        /// Called from Excel. The Application property(Excel.Application) from the addin, and task panes if exists, was automaticly disposed after these call
        /// </summary>
        /// <param name="RemoveMode"></param>
        /// <param name="custom"></param>
        void ThisAddIn_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }

        // this error handler is used for IExtensibility2 methods (your code) and the COMAddin methods GetCustomUI and CTPFactoryAvailable        
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            
        }
    }
}
