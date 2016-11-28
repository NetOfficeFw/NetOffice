using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Create a time stamp in host application addin key while register
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class TimestampAttribute : System.Attribute
    {
        /// <summary>
        /// Create a time stamp in host application addin key
        /// </summary>
        public TimestampAttribute()
        {

        }
    }
}
