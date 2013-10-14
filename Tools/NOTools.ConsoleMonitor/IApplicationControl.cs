using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    public interface IApplicationControl
    {
        string AddNewMessage(string notifyTime, string consoleChannelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain);
       
        void UpdateDisplayContent(bool showTime, bool showMachine, bool showAppDomain);
        void Clear();
        IApplicationHost Host { get; }

        string ControlName { get; }
    }
}
