using System;
namespace LateBindingApi.Core
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
        /// <param name="name"></param>
        /// <param name="versions"></param>
        public SupportByVersionAttribute(string name, params double[] versions)
        {
            this.Name = name;
            this.Versions = versions;
        }
    }
}
