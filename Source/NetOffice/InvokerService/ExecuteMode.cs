using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.InvokerService
{
    /// <summary>
    /// Invoker Service Execution Mode
    /// </summary>
    public enum ExecuteMode
    {
        /// <summary>
        /// Get Property
        /// </summary>
        PropertyGet = 0,

        /// <summary>
        /// Set Property
        /// </summary>
        PropertySet = 1,

        /// <summary>
        /// Method without return value
        /// </summary>
        Method = 2,

        /// <summary>
        /// Method with return value
        /// </summary>
        MethodGet = 3
    }
}
