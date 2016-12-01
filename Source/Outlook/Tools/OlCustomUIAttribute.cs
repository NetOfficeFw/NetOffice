using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Specify an embedded XML File for RibbonUI on a specific window
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = true)]
    public class OlCustomUIAttribute : System.Attribute
    {
        /// <summary>
        /// Explorer Ribbon ID
        /// </summary>
        private static string _mainWindowRibbonID = "Microsoft.Outlook.Explorer";

        /// <summary>
        /// Target Window ID
        /// </summary>
        public readonly string RibbonID;

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
        public OlCustomUIAttribute(string value)
        {
            RibbonID = _mainWindowRibbonID;
            Value = value;
        }
        
        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="ribbonID">Target Window ID</param>
        /// <param name="value">Full qualified location</param>
        public OlCustomUIAttribute(string ribbonID, string value)
        {
            if (String.IsNullOrEmpty(ribbonID))
                ribbonID = _mainWindowRibbonID;
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("value");

            RibbonID = ribbonID;
            Value = value;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="ribbonID">Target Window ID</param>
        /// <param name="value">Full qualified location</param>
        /// <param name="useAssemblyNamespace">Use namespace of the calling instance</param>
        public OlCustomUIAttribute(string ribbonID, string value, bool useAssemblyNamespace)
        {
            if (String.IsNullOrEmpty(ribbonID))
                throw new ArgumentException("ribbonID");
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("value");

            RibbonID = ribbonID;
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
