using NetOffice.CoreServices.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides access to cache services
    /// </summary>
    public interface ICoreCache
    {
        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }

        /// <summary>
        /// InterfaceId,ComponentId Cache
        /// </summary>
        /// <returns></returns>
        KeyValuePair<Guid, Guid>[] GetTypeComponentIdCache();

        /// <summary>
        /// Proxy, Contract, Implementation Cache
        /// </summary>
        /// <returns></returns>
        IEnumerable<TypeInformation> GetTypeCache();
    }
}
