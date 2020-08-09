using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice
{
    /// <summary>
    /// Handle access to shared underlying COM proxy. Not intended to be used outside of NetOffice infrastructure.
    /// </summary>
    public interface ICOMProxyShareProvider
    {
        /// <summary>
        /// Returns the inner proxy shared access handler.
        /// </summary>
        /// <returns>shared proxy</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        COMProxyShare GetProxyShare();

        /// <summary>
        /// Set the inner proxy shared access handler.
        /// The method wants to acquire the shared object one time.
        /// </summary>
        /// <param name="share">shared proxy</param>
        /// <exception cref="ArgumentNullException">Throws when given share is null(Nothing in Visual Basic)</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        void SetProxyShare(COMProxyShare share);
    }
}