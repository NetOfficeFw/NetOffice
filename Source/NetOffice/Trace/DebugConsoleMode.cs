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

        /// <summary>
        /// Hold all debug and exceptions logs in a internal string list
        /// </summary>
        MemoryList = 3,

        /// <summary>
        /// Debug log is redirected to System.Diagnostics.Trace
        /// </summary>
        Trace = 4
    }
}