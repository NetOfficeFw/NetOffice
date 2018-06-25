using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace NetOffice
{
    /// <summary>
    /// Indicates where an instance comes from.
    /// </summary>
    public interface ICOMObjectProxyService
    {
        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        /// <return>true if instance is create from a given proxy, false if instance is created from scratch</return>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        bool FromProxyService { get; }
    }
}
