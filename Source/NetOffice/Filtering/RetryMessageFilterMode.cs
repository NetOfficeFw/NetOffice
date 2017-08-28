using System;

namespace NetOffice.Filtering
{
    /// <summary>
    /// Specify the filter for an RetryMessageFilter instance
    /// </summary>
    public enum RetryMessageFilterMode
    {
        /// <summary>
        /// Try rejected call again immediately
        /// </summary>
        Immediately = 0,

        /// <summary>
        /// Try rejected call again after few milliseconds
        /// </summary>
        Delayed = 1,

        /// <summary>
        /// Dont try rejected call again
        /// </summary>
        None = 2
    }
}
