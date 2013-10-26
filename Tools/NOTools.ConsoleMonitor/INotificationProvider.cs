using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Income MessageHandler
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
    public delegate string UpdateMonitorInvoker(string notifyTime,
                                                string consoleChannelName,
                                                string machineName, 
                                                string appDomainFriendlyName, 
                                                string parentEntryID, 
                                                string message, 
                                                bool showTime,
                                                bool showMachine,
                                                bool showAppDomain);
    /// <summary>
    /// Interface for a Notification Listener
    /// </summary>
    public interface INotificationProvider : IDisposable
    {
        /// <summary>
        /// Start the Notification Listener
        /// </summary>
        void Start();

        /// <summary>
        /// Stop the Notification Listener
        /// </summary>
        void Stop();

        /// <summary>
        /// Notification Listener is currently enabled 
        /// </summary>
        bool IsRunning { get; }

        /// <summary>
        /// Incoming Console Notification
        /// </summary>
        event UpdateMonitorInvoker ConsoleNotification;

        /// <summary>
        /// Incoming Channel Notification
        /// </summary>
        event UpdateMonitorInvoker ChannelNotification;
    }
}
