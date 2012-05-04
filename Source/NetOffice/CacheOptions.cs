using System;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Cache options for the Factory->Initialize method
    /// </summary>
    public enum CacheOptions
    {
        /// <summary>
        /// clear current infos about existing types and and loaded NetOffice assemblies
        /// </summary>
        ClearExistingCache = 0,

        /// <summary>
        /// any new infos in Initialize was added to the existing cache
        /// </summary>
        KeepExistingCacheAlive = 1
    }
}
