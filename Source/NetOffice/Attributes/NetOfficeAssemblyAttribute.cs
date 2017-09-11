using System;
using System.Reflection;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates assembly is a NetOffice api assembly
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly)]
    public class NetOfficeAssemblyAttribute : System.Attribute
    {
        /// <summary>
        /// Full qualified type name
        /// </summary>
        public static readonly string FullName = "NetOffice.Attributes.NetOfficeAssemblyAttribute";

        /// <summary>
        /// Multiple version string
        /// </summary>
        public readonly string SupportedApiVersion;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="supportedApiVersion">multiple version string</param>
        public NetOfficeAssemblyAttribute(string supportedApiVersion)
        {
            this.SupportedApiVersion = supportedApiVersion;
        }

        /// <summary>
        /// Returns information an assembly is marked with NetOfficeAssemblyAttribute
        /// </summary>
        /// <param name="assembly">given assembly as any</param>
        /// <returns>true if attribute exists, otherwise false</returns>
        public static bool ContainsAttribute(Assembly assembly)
        {
            return assembly.GetCustomAttributes(typeof(NetOfficeAssemblyAttribute), true).Length > 0;
        }

        /// <summary>
        /// Returns supported api version trough NetOfficeAssemblyAttribute or null if not exists
        /// </summary>
        /// <param name="assembly">given assembly as any</param>
        /// <returns>supported api version or null</returns>
        public static string GetSupportedApiVersion(Assembly assembly)
        {
            object[] attributes = assembly.GetCustomAttributes(typeof(NetOfficeAssemblyAttribute), true);
            if (attributes.Length > 0)
                return (attributes[0] as NetOfficeAssemblyAttribute).SupportedApiVersion;
            else
                return null;
        }
    }
}