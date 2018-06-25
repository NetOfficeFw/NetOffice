using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates the native interface has an default wrapper implementation
    /// </summary>
    [AttributeUsage(AttributeTargets.Interface)]
    public class NativeCallerWrapperAttribute : System.Attribute
    {
        /// <summary>
        /// Wrapper Type
        /// </summary>
        public readonly Type Caller;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="caller">wrapper type</param>
        public NativeCallerWrapperAttribute(Type caller)
        {
            Caller = caller;
        }
    }
}
