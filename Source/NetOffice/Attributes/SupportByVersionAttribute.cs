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
        /// Name of the Library
        /// </summary>
        public readonly string Name;

        /// <summary>
        ///All supported library versions from the entity
        /// </summary>
        public readonly double[] Versions;

        /// <summary>
        /// Creates an instance of the attribute
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
