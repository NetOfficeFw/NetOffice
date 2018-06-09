using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Specify the origin COM type id
    /// </summary>
    [AttributeUsage(AttributeTargets.Interface)]
    public class TypeIdAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value"></param>
        /// <exception cref="ArgumentNullException">value is null</exception>
        /// <exception cref="FormatException"> value is not in a recognized format</exception>
        public TypeIdAttribute(string value)
        {
            Value = Guid.Parse(value);
        }

        /// <summary>
        /// Origin COM Type ID
        /// </summary>
        public readonly Guid Value;
    }
}
