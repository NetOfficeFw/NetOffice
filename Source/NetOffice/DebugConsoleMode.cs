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
        /// debug log are not enabled
        /// </summary>
        None = 0,

        /// <summary>
        /// debug log was redirected to System.Console
        /// </summary>
        Console = 1,

        /// <summary>
        /// debug log append to a logfile
        /// </summary>
        LogFile = 2,

        /// <summary>
        /// hold all debug and exceptions logs in a internal string list
        /// </summary>
        MemoryList = 3,

        /// <summary>
        /// debug log was redirected to System.Diagnostics.Trace
        /// </summary>
        Trace = 4
    }
}