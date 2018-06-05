using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides ICOMObject Activation Services
    /// </summary>
    public interface ICoreActivator
    {
        /// <summary>
        /// Occours when a new COMObject instance has been created
        /// </summary>
        event OnCreateInstanceEventHandler CreateInstance;

        /// <summary>
        /// Occurs when a new COMDynamicObject instance should be created
        /// </summary>
        event OnCreateCOMDynamicEventHandler CreateDynamicInstance;

        /// <summary>
        /// Occurs when a new COMProxyShare instance should be created
        /// </summary>
        event OnCreateProxyShareEventHandler CreateProxyShare;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }
    }
}
