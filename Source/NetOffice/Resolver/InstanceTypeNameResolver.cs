using System;

namespace NetOffice.Resolver
{
    /// <summary>
    /// Spend informations about ICOMObject instances
    /// </summary>
    internal class InstanceTypeNameResolver
    {
        /// <summary>
        /// Return the component root namespace of an ICOMObject instance
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>root namespace</returns>
        internal string GetComponentName(ICOMObject instance)
        {
            return instance.InstanceType.Namespace;
        }

        /// <summary>
        /// Return a human readable office-like instance type description of an ICOMObject instance
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>type description</returns>
        internal string GetFriendlyInstanceName(ICOMObject instance)
        {
            string result = instance.InstanceType.FullName;
            result = result.Replace("NetOffice.", String.Empty);
            result = result.Replace("Api", String.Empty);
            return result;
        }
    }
}
