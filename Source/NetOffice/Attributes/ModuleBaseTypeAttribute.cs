using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates a entity is a static helper module
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class ModuleBaseTypeAttribute : System.Attribute
    {
        /// <summary>
        /// Module Base Type
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">module base type</param>
        public ModuleBaseTypeAttribute(Type value)
        {
            Value = value;
        }
    }
}
