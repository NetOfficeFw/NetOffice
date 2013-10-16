using System;
namespace NetOffice
{
    /// <summary>
    /// Indicates which COM Type Library Version supports the entity
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class SupportByVersionAttribute : System.Attribute
    {
        /// <summary>
        /// returns library name
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// returns all supported library versions of entity
        /// </summary>
        public readonly double[] Versions;

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="name">name of the library</param>
        /// <param name="versions">version of the library</param>
        public SupportByVersionAttribute(string name, params double[] versions)
        {
            this.Name = name;
            this.Versions = versions;
        }
    }
}
