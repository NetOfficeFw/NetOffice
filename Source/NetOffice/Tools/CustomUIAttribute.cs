using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify an embedded XML File for RibbonUI
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class CustomUIAttribute : System.Attribute
    {
        /// <summary>
        /// Full qualified location
        /// </summary>
        public readonly string Value;

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="value">Full qualified location</param>
        public CustomUIAttribute(string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("value");

            Value = value;
        }
    }
}
