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

        /// <summary>
        /// Determines the given argument is a progid witout specific version
        /// </summary>
        /// <param name="value">progid as any</param>
        /// <returns>true if valid, otherwise false</returns>
        internal static bool ValidNonVersionedSignature(string value)
        {
            string[] result = new string[0];

            if(!String.IsNullOrWhiteSpace(value))
                result = value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);

            return result.Length == 2;
        }

        /// <summary>
        /// Returns the component from a prog id
        /// </summary>
        /// <param name="value">progid as any</param>
        /// <returns>component or empty if invalid progid</returns>
        internal static string Component(string value)
        {
            string[] result = new string[0];

            if (!String.IsNullOrWhiteSpace(value))
                 result = value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);

            return result.Length > 0 ? result[0] : String.Empty;
        }

        /// <summary>
        /// Returns the type from a prog id
        /// </summary>
        /// <param name="value">progid as any</param>
        /// <returns>type or empty if invalid progid</returns>
        internal static string Type(string value)
        {
            string[] result = new string[0];

            if (!String.IsNullOrWhiteSpace(value))
                result = value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);

            return result.Length > 1? result[1] : String.Empty;
        }
    }
}
