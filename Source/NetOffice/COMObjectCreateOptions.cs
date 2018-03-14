using System;

namespace NetOffice
{
    /// <summary>
    /// COMObject Create Flags
    /// </summary>
    public enum COMObjectCreateOptions
    {
        /// <summary>
        /// Nothing Special
        /// </summary>
        None = 0,

        /// <summary>
        /// Create and use own core instead of default core
        /// </summary>
        CreateNewCore = 1
    }

}