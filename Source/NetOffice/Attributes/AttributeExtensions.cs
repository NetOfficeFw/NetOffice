using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Spend generic attribute reflection extensions
    /// </summary>
    public static class AttributeExtensions
    {
        /// <summary>
        /// Determines an attribute exists on target type
        /// </summary>
        /// <typeparam name="T">target attribute type</typeparam>
        /// <param name="value">value to reflect</param>
        /// <param name="inherit">search also base types</param>
        /// <returns>true if attribute exists, otherwise false</returns>
        public static bool HasCustomAttribute<T>(this Type value, bool inherit = false) where T : Attribute
        {
            object[] attributes = value.GetCustomAttributes(typeof(T), inherit);
            return attributes.Length > 0;
        }

        /// <summary>
        /// Returns first found attribute or null
        /// </summary>
        /// <typeparam name="T">target attribute type</typeparam>
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

        /// <summary>
        /// Returns all found attributes 
        /// </summary>
        /// <typeparam name="T">target attribute type</typeparam>
        /// <param name="value">value to reflect</param>
        /// <param name="inherit">search also base types</param>
        /// <returns>target or null</returns>
        public static T[] GetCustomAttributes<T>(this Type value, bool inherit = false) where T : Attribute
        {
            T[] result = null;
            object[] attributes = value.GetCustomAttributes(typeof(T), inherit);
            result = new T[attributes.Length];
            for (int i = 0; i < attributes.Length; i++)
            {
                result[i] = attributes[i] as T;
            }
            return result;
        }
    }
}
