using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// ProgId to create an instance from
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class ComProgIdAttribute : System.Attribute
    {
        /// <summary>
        /// Registered ProgId if installed
        /// </summary>
        public readonly string Value;

        /// <summary>
        /// Creates an instance
        /// </summary>
        /// <param name="value">registered progId if installed</param>
        public ComProgIdAttribute(string value)
        {
            Value = value;
        }
    }
}
