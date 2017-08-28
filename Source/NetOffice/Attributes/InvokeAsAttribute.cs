using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates how an entity want invoke internally
    /// </summary>
    public enum Invoke
    {
        /// <summary>
        /// Invoke as property
        /// </summary>
        Property = 0,

        /// <summary>
        /// Invoke as method
        /// </summary>
        Method = 1
    }

    /// <summary>
    /// Invoke Internal usage
    /// </summary>
    public class InvokeAsAttribute : Attribute
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="invoke">invoke kind</param>
        public InvokeAsAttribute(Invoke invoke)
        {
            Invoke = invoke;
        }

        /// <summary>
        /// Invoke Kind
        /// </summary>
        public readonly Invoke Invoke;
    }
}
