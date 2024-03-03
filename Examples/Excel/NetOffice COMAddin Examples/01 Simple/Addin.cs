using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;
/*
    Minimum Addin Example
*/
namespace Excel01AddinCS4
{
    [COMAddin("Excel01 Sample Addin CS4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Excel01AddinCS4.Connect"), Guid("BB5D9F5A-267A-462E-9980-C65204969BE3"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine("Excel Version is {0}", Application.Version);   
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            
        }
    }
}