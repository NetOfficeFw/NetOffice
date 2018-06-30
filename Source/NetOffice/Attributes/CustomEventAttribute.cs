using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Custom NetOffice event and not provided from Office Automation Model
    /// </summary>
    [AttributeUsage(AttributeTargets.Event)]
    public class CustomEventAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public CustomEventAttribute()
        {

        }
    }
}
