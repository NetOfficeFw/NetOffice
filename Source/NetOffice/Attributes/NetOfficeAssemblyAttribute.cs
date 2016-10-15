using System;

namespace NetOffice
{
    /// <summary>
    /// Indicates assembly is a NetOffice api assembly
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly)]
    public class NetOfficeAssemblyAttribute : System.Attribute
    {
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
    }
}