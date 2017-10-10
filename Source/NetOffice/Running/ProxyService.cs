using System;
using System.Linq;
using System.Collections.Generic;    
using NetOffice.CollectionsGeneric;
using NetOffice.Contribution.CollectionsGeneric;

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
        /// <param name="componentName"></param>
        /// <param name="className"></param>
        /// <param name="throwExceptionIfNotFound"></param>
        /// <returns></returns>
        public static T GetActiveInstance<T>(string componentName, string className, bool throwExceptionIfNotFound = false) where T : class, ICOMObject
        {
            IDisposableSequence<T> result = GetActiveInstances<T>(componentName, className);
            T item = result.FirstOrDefault();
            result.Dispose(item);
            if (throwExceptionIfNotFound && null == item)
                throw new ArgumentOutOfRangeException(componentName + ", " + className);
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
        /// <returns>proxy instance or null</returns>
        public static object GetActiveInstance(string componentName, string className, bool throwExceptionIfNothingFound)
        {
            string compName = ValidateArgumentString(componentName);
            string clsName = ValidateArgumentString(className);

            if (compName == "EXCEL" && className == "APPLICATION")
            {
                object result = GetActiveExcelApplicationProxyFromDesktop();
                if (null == result && throwExceptionIfNothingFound)
                    throw new System.Runtime.InteropServices.COMException("Target instance is not running.");
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
