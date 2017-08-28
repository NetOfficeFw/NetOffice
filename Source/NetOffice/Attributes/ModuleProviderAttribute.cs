using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates the entity provides its underlying object to a static helper module
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class ModuleProviderAttribute : System.Attribute
    {
        /// <summary>
        /// Module Type
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">module type</param>
        public ModuleProviderAttribute(Type value)
        {
            Value = value;
        }
    }
}
