using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace NetOffice
{
    /// <summary>
    /// Indicates where the an instance comes from
    /// </summary>
    public interface ICOMObjectProxyService
    {
        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        bool FromProxyService { get; }
    }
}
