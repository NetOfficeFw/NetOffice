using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// DebugConsole operation mode 
    /// </summary>
    public enum DebugConsoleMode
    {
        /// <summary>
        /// Debug log is disabled
        /// </summary>
        None = 0,

        /// <summary>
        /// Debug log is redirected to System.Console
        /// </summary>
        Console = 1,

        /// <summary>
        /// Debug log append to a logfile
        /// </summary>
        LogFile = 2,

        /*
          MemoryList has been removed in NetOffice 1.7.4
          All messages goes automaticaly to the internal list now, regardless from the mode.
          Morevoer the message list want contains only 1000 items and remove the oldest automaticaly.
        */

        /// <summary>
        /// Debug log is redirected to System.Diagnostics.Trace
        /// </summary>
        Trace = 4
    }
}