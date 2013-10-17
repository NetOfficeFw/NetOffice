using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Host Application
    /// </summary>
    public interface IApplicationHost
    {
        /// <summary>
        /// An Application control is currently visible
        /// </summary>
        /// <param name="control">target control</param>
        /// <returns>true if visible</returns>
        bool IsCurrentlyVisible(IApplicationControl control);

        /// <summary>
        /// show time in messages
        /// </summary>
        bool ShowTime { get; }

        /// <summary>
        /// show machine in messages
        /// </summary>
        bool ShowMachine { get; }

        /// <summary>
        /// show appdomain in mesages
        /// </summary>
        bool ShowAppDomain { get; }
    }
}
