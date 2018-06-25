using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Callers
{
    /// <summary>
    /// Determines activation mode for InteropCompatibility classes
    /// </summary>
    public enum InteropCompatibilityClassCreateMode
    {
        /// <summary>
        /// Creates a new underyling proxy based on default prog id
        /// </summary>
        Direct = 0,

        /// <summary>
        /// Do nothing in order to wait for ICOMObjectInitialize call 
        /// </summary>
        FromActivator = 1
    }
}
