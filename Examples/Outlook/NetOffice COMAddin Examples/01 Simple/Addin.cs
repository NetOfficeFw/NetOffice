using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.OutlookApi.Tools;
/*
    Minimum Addin Example
*/
namespace Outlook01AddinCS4
{
    [COMAddin("Outlook01 Sample Addin CS4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Outlook01AddinCS4.Connect"), Guid("DF8FE853-A469-456C-94A2-BEA9CB735DEF"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine("Outlook Version is {0}", Application.Version);            
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            
        }
    }
}