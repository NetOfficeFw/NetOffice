using System;
using NetOffice;
using NetOffice.Exceptions;

namespace NetOffice.Extensions
{
    /// <summary>
    /// Caller extensions to write so called 'modern' code
    /// </summary>
    public static class DomainExtensions
    {
        /// <summary>
        /// Set a property value on ICOMObject instance and returns the instance
        /// </summary>
        /// <typeparam name="T">instance target result type</typeparam>
        /// <param name="value">ICOMObject instance</param>
        /// <param name="propertyName">target property name</param>
        /// <param name="propertyValue">target propertyValue</param>
        /// <returns>ICOMObject instance</returns>
        /// <exception cref="InvalidCastException">unable to cast calling instance to generic argument type or calling instance is not a com proxy</exception>
        /// <exception cref="NetOfficeCOMException">an unexpected error occurs on the remote server</exception>
        public static T SetProperty<T>(this object value, string propertyName, object propertyValue) where T : class
        {
            ICOMObject comObject = (ICOMObject)value;
            T result = (T)comObject;
            comObject.Factory.ExecuteValuePropertySet(comObject, propertyName, propertyValue);
            return result;
        }

        /// <summary>
        /// Calls member method on ICOMObject instance and returns the instance
        /// </summary>
        /// <typeparam name="T">instance target result type</typeparam>
        /// <param name="value">ICOMObject instance</param>
        /// <param name="methodName">target method name</param>
        /// <returns>ICOMObject instance</returns>
        /// <exception cref="InvalidCastException">unable to cast calling instance to generic argument type or calling instance is not a com proxy</exception>
        /// <exception cref="NetOfficeCOMException">an unexpected error occurs on the remote server</exception>
        public static T CallMethod<T>(this object value, string methodName) where T : class
        {
            ICOMObject comObject = (ICOMObject)value;
            T result = (T)comObject;
            comObject.Factory.ExecuteMethod(comObject, methodName);
            return result;
        }

        /// <summary>
        /// Calls member method on ICOMObject instance and returns the instance
        /// </summary>
        /// <typeparam name="T">instance target result type</typeparam>
        /// <param name="value">ICOMObject instance</param>
        /// <param name="methodName">target method name</param>
        /// <param name="values">arguments as any</param>
        /// <returns>ICOMObject instance</returns>
        /// <exception cref="InvalidCastException">unable to cast calling instance to generic argument type or calling instance is not a com proxy</exception>
        /// <exception cref="NetOfficeCOMException">an unexpected error occurs on the remote server</exception>
        public static T CallMethod<T>(this object value, string methodName, params object[] values) where T : class
        {
            ICOMObject comObject = (ICOMObject)value;
            T result = (T)comObject;
            comObject.Factory.ExecuteMethod(comObject, methodName, values);
            return result;
        }

        /// <summary>
        /// Calls member method on ICOMObject instance and returns the result as given generic argument type
        /// </summary>
        /// <typeparam name="T">target result type</typeparam>
        /// <param name="value">ICOMObject instance</param>
        /// <param name="methodName">target method name</param>
        /// <returns>result instance</returns>
        /// <exception cref="InvalidCastException">unable to cast calling instance to generic argument type or calling instance is not a com proxy</exception>
        /// <exception cref="NetOfficeCOMException">an unexpected error occurs on the remote server</exception>
        public static T CallMethodSelectResult<T>(this object value, string methodName)
        {
            ICOMObject comObject = (ICOMObject)value;
            T result = (T)comObject.Factory.ExecuteObjectMethodGet(comObject, methodName);
            return result;
        }

        /// <summary>
        /// Calls member method on ICOMObject instance and returns the result as given generic argument type
        /// </summary>
        /// <typeparam name="T">target result type</typeparam>
        /// <param name="value">ICOMObject instance</param>
        /// <param name="methodName">target method name</param>
        /// <param name="values">arguments as any</param>
        /// <returns>result instance</returns>
        /// <exception cref="InvalidCastException">unable to cast calling instance to generic argument type or calling instance is not a com proxy</exception>
        /// <exception cref="NetOfficeCOMException">an unexpected error occurs on the remote server</exception>
        public static T CallMethodSelectResult<T>(this object value, string methodName, params object[] values)
        {
            ICOMObject comObject = (ICOMObject)value;
            T result = (T)comObject.Factory.ExecuteObjectMethodGet(comObject, methodName, values);
            return result;
        }
    }
}