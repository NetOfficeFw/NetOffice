using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates the class is a default wrapper for a native interface caller
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class IsNativeCallerWrapperAttribute : System.Attribute
    {
        /// <summary>
        /// Native Caller
        /// </summary>
        public readonly Type Target;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="target">native caller</param>
        public IsNativeCallerWrapperAttribute(Type target)
        {
            Target = target;
        }
    }
}
