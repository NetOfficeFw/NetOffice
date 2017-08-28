using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice
{
    /// <summary>
    /// Handle access to share an underlying COM proxy. Not intended to use outside from Netoffice infrastructure.
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
        /// The method want aquire the share 1x times.
        /// </summary>
        /// <param name="share">shared proxy</param>
        /// <exception cref="ArgumentNullException">Throws when given share is null(Nothing in Visual Basic)</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        void SetProxyShare(COMProxyShare share);
    }
}