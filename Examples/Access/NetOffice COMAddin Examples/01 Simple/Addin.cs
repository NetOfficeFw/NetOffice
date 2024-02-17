using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.AccessApi.Tools;
/*
    Minimum Addin Example
*/
namespace Access01AddinCS4
{
    [COMAddin("Access01 Sample Addin CS4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Access01AddinCS4.Connect"), Guid("C1C784CA-751F-429D-AB39-ECAA7D38BD4D"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            if(Application.EntityIsAvailable("Version"))
                Console.WriteLine("Access Version is {0}", Application.Version);
            else
                Console.WriteLine("Access Version is {0}", "9(2000) or below");
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            
        }
    }
}