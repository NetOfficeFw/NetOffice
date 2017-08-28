using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Internal invoke member name, regardless of its wrapper name
    /// </summary>
    public class InternalNameAttribute : System.Attribute
    {
        /// <summary>
        /// The internal member name
        /// </summary>
        public readonly string InternalName;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="internalName">internal member name</param>
        public InternalNameAttribute(string internalName)
        {
            InternalName = internalName;
        }
    }
}
