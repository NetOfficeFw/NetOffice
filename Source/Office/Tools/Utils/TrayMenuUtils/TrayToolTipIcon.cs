using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Defines a set of standardized icons that can be associated with a ToolTip.
    /// </summary>
    public enum TrayToolTipIcon
    {
        /// <summary>
        ///  Not a standard icon.
        /// </summary>
        None = 0,

        /// <summary>
        /// An information icon.
        /// </summary>
        Info = 1,

        /// <summary>
        /// An information icon.
        /// </summary>
        Warning = 2,

        /// <summary>
        /// An error icon
        /// </summary>
        Error = 3
    }
}
