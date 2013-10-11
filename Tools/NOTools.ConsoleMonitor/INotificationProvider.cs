using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="notifyTime"></param>
    /// <param name="consoleChannelName"></param>
    /// <param name="appDomainFriendlyName"></param>
    /// <param name="parentEntryID"></param>
    /// <param name="message"></param>
    public delegate string UpdateMonitorInvoker(string notifyTime,
                                                string consoleChannelName,
                                                string machineName, 
                                                string appDomainFriendlyName, 
                                                string parentEntryID, 
                                                string name, 
                                                bool showTime,
                                                bool showMachine,
                                                bool showAppDomain);
    /// <summary>
    /// 
    /// </summary>
    public interface INotificationProvider : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        void Start();

        /// <summary>
        /// 
        /// </summary>
        void Stop();

        /// <summary>
        /// 
        /// </summary>
        bool IsRunning { get; }

        /// <summary>
        /// 
        /// </summary>
        event UpdateMonitorInvoker ConsoleNotification;

        /// <summary>
        /// 
        /// </summary>
        event UpdateMonitorInvoker ChannelNotification;
    }
}
