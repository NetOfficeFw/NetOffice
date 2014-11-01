using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// some helper methods (also for visual basic)
    /// </summary>
    public class Utils
    {
        /// <summary>
        /// lock instance to perform threadsafe operations
        /// </summary>
        private static object _lockUtils = new object();

        /// <summary>
        /// Shared Default Invoker
        /// </summary>
        public static Utils Default
        {
            get
            {
                lock (_lockUtils)
                {
                    if (null == _default)
                        _default = new Utils();
                    return _default;
                }
            }
        }
        private static Utils _default;

        /// <summary>
        /// checks value is null or nothing. 
        /// </summary>
        /// <param name="value">check value</param>
        /// <returns>true if null</returns>
        public static bool IsNothing(object value)
        {
            return null == value;
        }

        /// <summary>
        /// returns enumerator with com proxies
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetProxyEnumeratorAsProperty(COMObject comObject)
        {
            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                COMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true);
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    COMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                    yield return returnClass;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// returns enumerator with com proxies
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetProxyEnumeratorAsMethod(COMObject comObject)
        {
            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                COMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true);
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    COMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                    yield return returnClass;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// returns enumerator with scalar variables
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetScalarEnumeratorAsProperty(COMObject comObject)
        {
            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                COMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true);
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object item = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    yield return item;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// returns enumerator with scalar variables
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetScalarEnumeratorAsMethod(COMObject comObject)
        {
            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                COMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true);
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object item = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    yield return item;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

    }
}