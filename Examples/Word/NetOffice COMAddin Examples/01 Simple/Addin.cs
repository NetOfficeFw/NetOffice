using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.WordApi.Tools;
/*
    Minimum Addin Example
*/
namespace Word01AddinCS4
{
    [COMAddin("Word01 Sample Addin CS4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Word01AddinCS4.Connect"), Guid("D4282C02-B127-4FCA-835D-2CD86B0CB00A"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine("Word Version is {0}", Application.Version);            
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            
        }
    }
}