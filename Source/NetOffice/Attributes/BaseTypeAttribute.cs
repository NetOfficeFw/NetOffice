using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Well known base type. That means other types inherit from.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = false)]
    public class BaseTypeAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        public BaseTypeAttribute()
        {

        }
    }
}
