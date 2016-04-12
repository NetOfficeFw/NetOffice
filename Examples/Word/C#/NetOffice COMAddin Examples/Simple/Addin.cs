using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.WordApi.Tools;
using NetOffice.WordApi;

/*
  * This project shows you the COMAddin base class from the NetOffice tools.
  * Its designed to reduce infrastructure code from your own.
  * You have to set some attributes and thats all. 
  * You see also the host application is available as class instance property. no need for dispose here because the base class do this for you while shutdown.
*/

namespace NetOfficeTools.SimpleAccessCS4
{
    [COMAddin("NetOfficeCS4 Sample Word Addin", "This Addin shows you the COMAddin base class from the NetOffice Tools", 3)]
    [Guid("F4EC5A72-C431-4BFA-85AB-43A0817617C7"), ProgId("SimpleWordCS4.Addin")]
    public class Addin : COMAddin
    {
        public Addin()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            // show the host application version
            string hostVersion = String.Format("Host Application Version is:{0}", this.Application.Version);
            Utils.Dialog.ShowMessageBox(hostVersion, MessageBoxIcon.Information, DialogResult.OK);
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

        }
    }
}
