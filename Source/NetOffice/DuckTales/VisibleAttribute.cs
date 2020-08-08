using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Duck
{
    /// <summary>
    /// Indicates a special method, name starts with get_ or set_ which is not compiler generated
    /// </summary>
    [Obsolete("Support for dynamic objects will be removed in NetOffice 1.8")]
    public class VisibleAttribute : System.Attribute
    {
        /// <summary>
        /// Always true
        /// </summary>
        public readonly bool Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public VisibleAttribute()
        {
            Value = true;
        }
    }
}
