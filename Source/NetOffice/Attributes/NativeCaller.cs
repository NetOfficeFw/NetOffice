using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Instance use early bind calls for underlying object
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class NativeCallerAttribute : System.Attribute
    {
        /// <summary>
        /// Native Interface Type
        /// </summary>
        public readonly Type Native;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="native">native interface type</param>
        public NativeCallerAttribute(Type native)
        {
            Native = native;
        }
    }
}