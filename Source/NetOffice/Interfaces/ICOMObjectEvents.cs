using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Represents informations about the event state
    /// </summary>
    public interface ICOMObjectEvents
    {
        /// <summary>
        /// Returns information the instance offers events
        /// </summary>
        bool IsEventBinding { get; }

        /// <summary>
        /// Returns information the instance has been initialized the internal event bridge
        /// </summary>
        bool IsEventBridgeInitialized { get; }

        /// <summary>
        /// Returns the count of event subscriptions from the instance
        /// </summary>
        bool IsWithEventRecipients { get; }
    }
}
