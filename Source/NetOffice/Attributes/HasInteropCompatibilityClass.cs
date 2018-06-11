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
    public class HasInteropCompatibilityClassAttribute : System.Attribute
    {
        /// <summary>
        /// Interop Compatibility Class
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">interop compatibility class</param>
        public HasInteropCompatibilityClassAttribute(Type value)
        {
            Value = value;
        }
    }
}
