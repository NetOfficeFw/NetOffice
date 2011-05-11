using System;
namespace LateBindingApi.Core
{
    /// <summary>
    /// Indicates whitch COM Type Library Version supports the entity
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class SupportByLibrary : System.Attribute
    {
        /// <summary>
        /// returns all supported library versions of entity
        /// </summary>
        public readonly string[] Versions;

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="versions"></param>
        public SupportByLibrary(params string[] versions)
        {
            this.Versions = versions;
        }
    }
}
