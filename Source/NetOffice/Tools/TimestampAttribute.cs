using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Create a time stamp in host application add-in registry key when registering the add-in.
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class TimestampAttribute : System.Attribute
    {
        /// <summary>
        /// Create a time stamp in host application add-in registry key.
        /// </summary>
        public TimestampAttribute()
        {

        }
    }
}
