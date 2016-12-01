using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify an embedded XML File for RibbonUI
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class CustomUIAttribute : System.Attribute
    {
        /// <summary>
        /// Full qualified location
        /// </summary>
        public readonly string Value;

        /// <summary>
        /// Use root namespace of the calling instance
        /// </summary>
        public readonly bool UseAssemblyNamespace;

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

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="value">Full qualified location</param>
        /// <param name="useAssemblyNamespace">Use namespace of the calling instance</param>
        public CustomUIAttribute(string value, bool useAssemblyNamespace)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("value");

            Value = value;
            UseAssemblyNamespace = useAssemblyNamespace;
        }

        /// <summary>
        /// Build resource path with attribute values
        /// </summary>
        /// <param name="resourcePath">resource path</param>
        /// <param name="useAssemblyNamespace">use assembly namespace</param>
        /// <param name="assemblyNamespace">root assembly namespace</param>
        /// <returns>return resource path</returns>
        public static string BuildPath(string resourcePath, bool useAssemblyNamespace, string assemblyNamespace)
        {
            if (String.IsNullOrEmpty(resourcePath))
                throw new ArgumentNullException("resourcePath");
            if (String.IsNullOrEmpty(assemblyNamespace))
                throw new ArgumentNullException("assemblyNamespace");

            if (useAssemblyNamespace)
            {
                string result = assemblyNamespace + "." + resourcePath;
                result = result.Replace("..", ".");
                return result;
            }
            else
            {
                return resourcePath;
            }
        }
    }
}
