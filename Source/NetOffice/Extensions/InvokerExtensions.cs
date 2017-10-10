using NetOffice.Exceptions;
using System;

namespace NetOffice.Extensions.Invoker
{
    /// <summary>
    /// ICOMObject Invoker Extensions
    /// </summary>
    public static class InvokerExtensions
    {
        /// <summary>
        /// Invoke Property if instance implement ICOMObject
        /// </summary>
        /// <typeparam name="T">result type</typeparam>
        /// <param name="value">instance which is ICOMObject</param>
        /// <param name="propertyName">name of property</param>
        /// <returns>result of invoked property or default(T) if instance doesnt implement ICOMObject</returns>
        public static T Property<T>(this object value, string propertyName)
        {
            ICOMObject comObject = value as ICOMObject;
            if (null != comObject)
            {
                object result = CorePropertyGetExtensions.ExecuteVariantPropertyGet(comObject.Factory, comObject, propertyName);
                if (result is T)
                    return (T)result;
                else
                    return default(T);
            }
            else
                return default(T);
        }

        /// <summary>
        /// Invoke method if instance implement ICOMObject
        /// </summary>
        /// <param name="value">instance which is ICOMObject</param>
        /// <param name="methodName">name of method</param>
        /// <param name="throwException">throw exception if unable to cast ICOMObject</param>
        /// <returns>result of invoked property or default(T) if instance doesnt implement ICOMObject</returns>
        public static void Method(this object value, string methodName, bool throwException = false)
        {
            ICOMObject comObject = value as ICOMObject;
            if (null != comObject)
            {
                CoreMethodExtensions.ExecuteMethod(comObject.Factory, comObject, methodName);
            }
            else
            {
                if(throwException)
                    throw new MethodCOMException("Unable to cast ICOMObject.");
            }
        }

        /// <summary>
        /// Invoke method if instance implement ICOMObject
        /// </summary>
        /// <param name="value">instance which is ICOMObject</param>
        /// <param name="methodName">name of method</param>
        /// <param name="paramsArray">arguments as any</param>
        /// <param name="throwException">throw exception if unable to cast ICOMObject</param>
        /// <returns>result of invoked property or default(T) if instance doesnt implement ICOMObject</returns>
        public static T MethodGet<T>(this object value, string methodName, object[] paramsArray, bool throwException = false)
        {
            ICOMObject comObject = value as ICOMObject;
            if (null != comObject)
            {
                object result = CoreMethodExtensions.ExecuteVariantMethodGet(comObject.Factory, 
                                                comObject, methodName, paramsArray);
                if (result is T)
                    return (T)result;
                else
                    return default(T);
            }
            else
            {
                if (throwException)
                    throw new MethodCOMException("Unable to cast ICOMObject.");
                else
                     return default(T);
            }
        }
    }
}