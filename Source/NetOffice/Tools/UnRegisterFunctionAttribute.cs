using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Mark a static method as UnRegister method. the method need the following signature public void UnRegister(Type type, RegisterCall callType)
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method, AllowMultiple = false)]
    public class UnRegisterFunctionAttribute : System.Attribute
    {
        /// <summary>
        /// Register Call Condition
        /// </summary>
        public readonly RegisterMode Value;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="mode">register call condition</param>
        public UnRegisterFunctionAttribute(RegisterMode mode)
        {
            Value = mode;
        }
    }
}
