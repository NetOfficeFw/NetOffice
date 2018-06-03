using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates there is an interop compatibility class for this type
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class HasInteropCompatibilityClass : System.Attribute
    {
        /// <summary>
        /// Interop Compatibility Class
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// nterop compatibility class
        /// </summary>
        /// <param name="value"></param>
        public HasInteropCompatibilityClass(Type value)
        {
            Value = value;
        }
    }
}
