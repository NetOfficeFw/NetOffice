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
        /// version
        /// </summary>
        public readonly string SupportedApiVersion;

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="apiVersion"></param>
        public NetOfficeAssemblyAttribute(string apiVersion)
        {
            this.SupportedApiVersion = apiVersion;
        }
    }
}