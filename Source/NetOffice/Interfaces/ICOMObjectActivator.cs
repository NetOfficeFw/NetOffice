using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Provides ICOMObject Activation Services
    /// </summary>
    public interface ICOMObjectActivator
    {
        /// <summary>
        /// Occours when a new COMObject instance has been created
        /// </summary>
        event COMObjectActivator.OnCreateInstanceEventHandler CreateInstance;

        /// <summary>
        /// Occurs when a new COMDynamicObject instance should be created
        /// </summary>
        event COMObjectActivator.OnCreateCOMDynamicEventHandler CreateDynamicInstance;

        /// <summary>
        /// Occurs when a new COMProxyShare instance should be created
        /// </summary>
        event COMObjectActivator.OnCreateProxyShareEventHandler CreateProxyShare;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }
    }
}
