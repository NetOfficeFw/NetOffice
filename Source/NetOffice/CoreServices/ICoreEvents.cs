using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides access to NetOffice Events Sinks services
    /// </summary>
    public interface ICoreEvents
    {
        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }

        /// <summary>
        /// Count of current opened event bridges
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Dispose all open event bridges
        /// </summary>
        void DisposeAllEventBridges();
    }
}
