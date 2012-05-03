using System;

namespace NetOffice
{
    /// <summary>
    /// Indicates assembly is a latebinding api assembly
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly)]
    public class LateBindingAttribute: System.Attribute
    {
        /// <summary>
        /// version
        /// </summary>
        public readonly string SupportedApiVersion;

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="apiVersion"></param>
        public LateBindingAttribute(string apiVersion)
        {
            this.SupportedApiVersion = apiVersion;
        }
    }
}




