using System;
using System.Linq;
using System.Collections.Generic;
using NetOffice.CollectionsGeneric;
using NetOffice.Contribution.CollectionsGeneric;
using NetOffice.Exceptions;

namespace NetOffice.Running
{
    /// <summary>
    /// Try to find running com instances.
    /// The strategy to find is -depending on the arguments- the Running Object Table(ROT) or Windows Desktop Subsystem through IAccessible.
    /// The reason is because the Running Object Table wont give you all com instances - its just shows only the informations.
    /// </summary>
    public static class ProxyService
    {
        /// <summary>
        ///  Returns all running com proxies, wrapped by T
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        ///  <param name="predicate">filter predicate</param>
        /// <returns>ICOMObject enumerator</returns>
        public static IDisposableSequence<T> GetActiveInstances<T>(string componentName, string className, Func<T, bool> predicate) where T : class, ICOMObject
        {
            Type typeOfT = typeof(T);
            IDisposableSequence instances = GetActiveInstances(componentName, className);
            List<T> result = new List<T>();
            foreach (object item in instances)
            {
                T newItem = Activator.CreateInstance(typeOfT, new object[] { null, item }) as T;
                if (null != predicate && predicate(newItem))
                {
                    result.Add(newItem);
                }
            }
            return new DisposableGenericList<T>(result.ToArray());
        }

        /// <summary>
        ///  Returns all running com proxies, wrapped by T
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <returns>ICOMObject enumerator</returns>
        public static IDisposableSequence<T> GetActiveInstances<T>(string componentName, string className) where T : class, ICOMObject
        {
            Type typeOfT = typeof(T);
            IDisposableSequence instances = GetActiveInstances(componentName, className);
            List<T> result = new List<T>();
            foreach (object item in instances)
            {
                T newItem = Activator.CreateInstance(typeOfT, new object[] { null, item }) as T;
                result.Add(newItem);
            }
            return new DisposableGenericList<T>(result.ToArray());
        }

        /// <summary>
        ///  Returns first running com proxy, wrapped by T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <param name="throwExceptionIfNotFound">throw ArgumentOutOfRangeException if no instance match</param>
        /// <returns>target instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="NetOfficeCOMException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static T GetActiveInstance<T>(string componentName, string className, bool throwExceptionIfNotFound = false) where T : class, ICOMObject
        {
            IDisposableSequence<T> result = GetActiveInstances<T>(componentName, className);
            T item = result.FirstOrDefault();
            result.Dispose(item);
            if (throwExceptionIfNotFound && null == item)
                throw new NetOfficeCOMException(String.Format("Unable to find active instance {0}.", componentName + ", " + className));
            return item;
        }

        /// <summary>
        ///  Returns first running com proxy, wrapped by T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw ArgumentOutOfRangeException if no instance match</param>
        /// <returns>target instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="NetOfficeCOMException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static T GetActiveInstance<T>(string componentName, string className, Func<T, bool> predicate, bool throwExceptionIfNotFound = false) where T : class, ICOMObject
        {
            IDisposableSequence<T> result = GetActiveInstances<T>(componentName, className, predicate);
            T item = result.FirstOrDefault();
            result.Dispose(item);
            if (throwExceptionIfNotFound && null == item)
                throw new NetOfficeCOMException(String.Format("Unable to find active instance {0}.", componentName + ", " + className));
            return item;
        }

        /// <summary>
        /// Returns all running com proxies there is match with given arguments
        /// </summary>
        /// <param name="componentName">component name, for example Excel</param>
        /// <param name="className">class name, for example Application</param>
        /// <returns>COM proxy enumerator</returns>
        public static IDisposableSequence GetActiveInstances(string componentName, string className)
        {
            string compName = ValidateArgumentString(componentName);
            string clsName = ValidateArgumentString(className);

            if (compName == "EXCEL" && className == "APPLICATION")
            {
                return GetActiveExcelApplicationProxiesFromDesktop();
            }
            else
            {
                return RunningObjectTable.GetActiveProxies(componentName, className);
            }
        }

        /// <summary>
        /// Returns a running com proxy there is match with given arguments
        /// </summary>
        /// <param name="componentName">component name, for example Excel or null as wildcard</param>
        /// <param name="className">class name, for example Application or null as wildcard</param>
        /// <param name="throwExceptionIfNothingFound">throw an exception if no proxy was found</param>
        /// <returns>proxy instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="NetOfficeCOMException">no instance found and throwExceptionIfNothingFound is set</exception>
        public static object GetActiveInstance(string componentName, string className, bool throwExceptionIfNothingFound)
        {
            string compName = ValidateArgumentString(componentName);
            string clsName = ValidateArgumentString(className);

            if (compName == "EXCEL" && className == "APPLICATION")
            {
                object result = GetActiveExcelApplicationProxyFromDesktop();
                if (null == result && throwExceptionIfNothingFound)
                    throw new NetOfficeCOMException(String.Format("Unable to find active instance {0}.", componentName + ", " + className));
                return result;
            }
            else
            {
                return RunningObjectTable.GetActiveProxy(componentName, className, throwExceptionIfNothingFound);
            }
        }

        private static object GetActiveExcelApplicationProxyFromDesktop()
        {
            try
            {
                WindowEnumerator enumerator = new WindowEnumerator("XLMAIN");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);
                if (null == handles || handles.Length == 0)
                    return null;

                object proxy = ExcelApplicationWindow.GetApplicationProxyFromHandle(handles[0]);
                return proxy;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        private static IDisposableSequence GetActiveExcelApplicationProxiesFromDesktop()
        {
            try
            {
                WindowEnumerator enumerator = new WindowEnumerator("XLMAIN");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);
                if (null == handles || handles.Length == 0)
                    return new DisposableObjectList(null);

                return ExcelApplicationWindow.GetApplicationProxiesFromHandle(handles);
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        private static string ValidateArgumentString(string arg)
        {
            string result = arg;
            if (result == null)
                result = String.Empty;
            result = result.Trim().ToUpper();
            return result;
        }
    }
}
