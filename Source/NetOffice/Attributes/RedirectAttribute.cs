using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates the member is an alias and want call another instance member
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    public class RedirectAttribute : System.Attribute
    {
        /// <summary>
        /// Instance member to call
        /// </summary>
        public readonly string Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">instance member to call</param>
        public RedirectAttribute(string name)
        {
            Value = name;
        }
    }
}
