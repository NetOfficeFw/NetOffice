using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Determine the kind of internal entity
    /// </summary>
    public enum InternalEntityKind
    {
        /// <summary>
        /// Sink helper to bridge com event interface to CoClass wrapper
        /// </summary>
        SinkHelper = 0,

        /// <summary>
        /// COM Interop Event interface
        /// </summary>
        ComEventInterface = 1
    }

    /// <summary>
    /// Indicates the entity is not intended to use by client caller directly
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class InternalEntityAttribute : System.Attribute
    {
        /// <summary>
        /// Entity Kind
        /// </summary>
        public readonly InternalEntityKind Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">entity kind</param>
        public InternalEntityAttribute(InternalEntityKind value)
        {
            Value = value;
        }
    }
}
