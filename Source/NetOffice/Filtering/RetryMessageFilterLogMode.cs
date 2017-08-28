using System;

namespace NetOffice.Filtering
{
    /// <summary>
    /// Specify log behaviour for an RetryMessageFilter instance
    /// </summary>
    public enum RetryMessageFilterLogMode
    {
        /// <summary>
        /// Disable Log
        /// </summary>
        None = 0,

        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.RetryRejectedCall
        /// </summary>
        RetryRejectedCall = 1,

        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.MessagePending
        /// </summary>
        MessagePending = 2,

        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.RetryRejectedCall and IMessageFilter.MessagePending
        /// </summary>
        Both = 3
    }
}
