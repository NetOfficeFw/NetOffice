using System;

namespace NetOffice
{
    /// <summary>
    /// Cache options for the <see cref="Core.Initialize()"/> method.
    /// </summary>
    public enum CacheOptions
    {
        /// <summary>
        /// Clear current information about existing types and loaded NetOffice assemblies.
        /// </summary>
        ClearExistingCache = 0,

        /// <summary>
        /// Any new type or NetOffice assembly information will be added to the existing cache.
        /// </summary>
        KeepExistingCacheAlive = 1
    }
}
