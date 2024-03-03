using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.PowerPointApi.Tools;
/*
    Minimum Addin Example
*/
namespace PowerPoint01AddinCS4
{
    [COMAddin("PowerPoint01 Sample Addin CS4", "Miminum Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("PowerPoint01AddinCS4.Connect"), Guid("C6AE4095-4F49-4A8F-9923-0F144C10E837"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
            OnStartupComplete += Addin_OnStartupComplete;
            OnDisconnection += Addin_OnDisconnection;
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine("PowerPoint Version is {0}", Application.Version);            
        }

        private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            
        }
    }
}