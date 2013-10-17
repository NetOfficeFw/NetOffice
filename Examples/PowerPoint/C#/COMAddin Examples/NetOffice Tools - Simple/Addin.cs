using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Tools;
using NetOffice.PowerPointApi;

/*
  * This project shows you the COMAddin base class from the NetOffice tools.
  * Its designed to reduce infrastructure code from your own.
  * You have to set some attributes and thats all. 
  * You see also the host application is available as class instance property. no need for dispose here because the base class do this for you while shutdown.
*/

namespace NetOfficeTools.SimplePPointCS4
{
    [COMAddin("NetOfficeCS4 Sample PowerPoint Addin", "This Addin shows you the COMAddin base class from the NetOffice Tools", 3)]
    [Guid("DD560ED5-C600-4D99-ADC6-9675511C3387"), ProgId("SimplePPointCS4.Addin"), Tweak(true)]
    public class Addin : COMAddin
    {
        public Addin()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
        }

        void Addin_OnStartupComplete(ref Array custom)
        {
            // get the host application version
            string hostVersion = this.Application.Version;
            Console.WriteLine("Host Application Version is:{0}", this.Application.Version);
        }

        void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

        }
    }
}
