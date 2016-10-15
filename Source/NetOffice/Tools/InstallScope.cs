using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{       
    /// <summary>
    /// The current install/uninstall scope
    /// </summary>
    [System.Runtime.InteropServices.Guid("FC5DC88D-D4D8-4BC8-A206-F55E7CD94C89")]
    public enum InstallScope
    {
        /// <summary>
        /// Whole Sstem
        /// </summary>
        System = 0,

        /// <summary>
        /// Current User
        /// </summary>
        User = 1
    }
}
