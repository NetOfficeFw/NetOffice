using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice
{
    /// <summary>
    /// Handle the shared access RCW behavior. Not intended to use from client callers.
    /// </summary>
    public interface ICOMProxyShareProvider
    {
        /// <summary>
        /// Returns the inner proxy shared access handler
        /// </summary>
        /// <returns>shared proxy</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        COMProxyShare GetProxyShare();

        /// <summary>
        /// Set the inner proxy shared access handler.
        /// The method want aquire the share 1x times
        /// </summary>
        /// <param name="share">target share</param>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        void SetProxyShare(COMProxyShare share);
    }
}

