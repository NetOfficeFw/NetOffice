using System;
using System.ComponentModel;

namespace NetOffice.Resolver
{
    /// <summary>
    /// Spend informations about the underlying proxy of an ICOMObject instance
    /// </summary>
    internal class UnderlyingTypeNameResolver
    {
        /// <summary>
        /// Returns the name of the hosting type library
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>library name or empty if its failed to recieve</returns>
        internal string GetComponentName(ICOMObject instance)
        {
            return null != instance ? TypeDescriptor.GetComponentName(instance.UnderlyingObject) : String.Empty;
        }

        /// <summary>
        /// Returns the class name of the underlying proxy
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>class name or empty if its failed to recieve</returns>
        internal string GetClassName(ICOMObject instance)
        {
            return null != instance ? TypeDescriptor.GetClassName(instance.UnderlyingObject) : String.Empty;
        }

        /// <summary>
        /// Returns a human readable office-like name of underlying proxy class
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="className">cached or default class name - want return if not null</param>
        /// <returns>friendly class name</returns>
        internal string GetFriendlyClassName(ICOMObject instance, string className)
        {
            string fullname = null != className ? className : GetFriendlyClassName(instance);
            return fullname;
        }

        /// <summary>
        /// Returns a human readable office-like name of underlying proxy class
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>friendly class name</returns>
        internal string GetFriendlyClassName(ICOMObject instance)
        {
            string fullName = instance.UnderlyingType.FullName;
            fullName = fullName.Replace("Microsoft", String.Empty).Replace("Interop", String.Empty).Replace("..", ".");
            return fullName;
        }
    }
}
