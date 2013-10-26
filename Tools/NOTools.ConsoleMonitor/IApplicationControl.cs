using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Defines a Application Control
    /// </summary>
    public interface IApplicationControl
    {
        /// <summary>
        /// add new log message to show
        /// </summary>
        /// <param name="notifyTime">sender time of notification</param>
        /// <param name="consoleChannelName">name of console oder channel</param>
        /// <param name="machineName">sender machine</param>
        /// <param name="appDomainFriendlyName">sender appdomain</param>
        /// <param name="parentEntryID">parent entry log id</param>
        /// <param name="name">message</param>
        /// <param name="showTime">display time</param>
        /// <param name="showMachine">display machine</param>
        /// <param name="showAppDomain">display appdomain</param>
        /// <returns>reated entry for the message</returns>
        string AddNewMessage(string notifyTime, string consoleChannelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain);

        /// <summary>
        /// Update/Refresh
        /// </summary>
        /// <param name="showTime">show time in messages</param>
        /// <param name="showMachine">show machine in messages</param>
        /// <param name="showAppDomain">show appdomain in messages</param>
        void UpdateDisplayContent(bool showTime, bool showMachine, bool showAppDomain);

        /// <summary>
        /// Clear Display Content
        /// </summary>
        void Clear();

        /// <summary>
        /// Parent Application
        /// </summary>
        IApplicationHost Host { get; }

        /// <summary>
        /// Unique name of the control
        /// </summary>
        string ControlName { get; }
    }
}
