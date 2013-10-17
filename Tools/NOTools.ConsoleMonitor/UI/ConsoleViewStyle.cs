using System;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// View style for ConsoleViewControl
    /// </summary>
    public enum ConsoleViewStyle
    {
        /// <summary>
        /// Top/Down. The last message is at bottom 
        /// </summary>
        Plain = 0,

        /// <summary>
        /// Down/Top. The last message is at top
        /// </summary>
        PlainReverse = 1,

        /// <summary>
        /// Top/Down with child hierarchy if support by client sender
        /// </summary>
        Hierarchy = 2
    }
}
