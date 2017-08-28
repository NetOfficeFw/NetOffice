using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates how an default property want invoke internally
    /// </summary>
    public enum IndexInvoke
    {
        /// <summary>
        /// Invoke as property
        /// </summary>
        Property = 0,

        /// <summary>
        /// Invoke as method
        /// </summary>
        Method = 1,
    }

    /// <summary>
    /// Default Property Implementation Details
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class HasIndexPropertyAttribute : System.Attribute
    {
        /// <summary>
        /// Internal Invoke Kind
        /// </summary>
        public readonly IndexInvoke Invoke;

        /// <summary>
        /// Internal Invoke Name
        /// </summary>
        public readonly string InvokeName;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="invoke">internal Invoke Kind</param>
        /// <param name="invokeName">internal Invoke Name</param>
        public HasIndexPropertyAttribute(IndexInvoke invoke, string invokeName)
        {
            Invoke = invoke;
            InvokeName = invokeName;
        }
    }

    /// <summary>
    /// Flags an default property
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IndexPropertyAttribute : System.Attribute
    {

    }
}
