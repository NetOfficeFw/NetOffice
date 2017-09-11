using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Spend generic attribute reflection extensions
    /// </summary>
    public static class AttributeExtensions
    {
        /// <summary>
        /// Returns first found attribute or null
        /// </summary>
        /// <typeparam name="T">target attriute</typeparam>
        /// <param name="value">value to reflect</param>
        /// <param name="inherit">search also base types</param>
        /// <returns>target or null</returns>
        public static T GetCustomAttribute<T>(this Type value, bool inherit = false) where T : Attribute
        {
            object[] attributes = value.GetCustomAttributes(typeof(T), inherit);
            if (attributes.Length > 0)
                return attributes[0] as T;
            else
                return null;
        }
    }
}