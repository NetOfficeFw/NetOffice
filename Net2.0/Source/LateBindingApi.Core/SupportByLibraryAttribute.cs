using System;
namespace LateBindingApi.Core
{
    /// <summary>
    /// Indicates whitch COM Type Library Version supports the entity
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class SupportByLibrary: System.Attribute
    {
        public readonly string[] Versions;

        public SupportByLibrary(params string[] versions)
        {
            this.Versions = versions;
        }
    }
}
