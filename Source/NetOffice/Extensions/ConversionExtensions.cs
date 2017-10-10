using System;
using NetOffice;
using NetOffice.Exceptions;

namespace NetOffice.Extensions.Conversion
{
    /// <summary>
    /// ICOMObject Conversion Extensions
    /// </summary>
    public static class ConversionExtensions
    {
        /// <summary>
        /// Cast instance to ICOMObject and clone instance as target type of T
        /// </summary>
        /// <typeparam name="T">given target type</typeparam>
        /// <param name="value">instance to convert</param>
        /// <param name="throwException">return null or throw exception if its failed to convert</param>
        /// <returns>instance of T or null(Nothing in Visual Basic)</returns>
        /// <exception cref="CloneException">Failed to convert instance to ICOMObject</exception>
        public static T To<T>(this object value, bool throwException = false) where T:class, ICOMObject
        {
            ICOMObject comObject = value as ICOMObject;
            if (null != comObject)
            {
                return comObject.To<T>();
            }
            else
            {
                if (throwException)
                    throw new CloneException("Unable to cast ICOMObject.");
                else
                    return null;
            }
        }
    }
}
