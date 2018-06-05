using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides access to the resolve services
    /// </summary>
    public interface ICoreResolver
    {
        /// <summary>
        /// Occurs when its failed to resolve a wrapper for a recieved com proxy.
        /// This event allows to find and set the corresponding wrapper at hand.
        /// Otherwise NetOffice want create a dynamic instance if possible.
        /// </summary>
        event ResolveEventHandler Resolve;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }
    }
}
