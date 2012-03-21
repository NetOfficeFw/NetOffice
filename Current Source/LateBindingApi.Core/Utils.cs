using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// some helper methods (also for visual basic)
    /// </summary>
    public static class Utils
    {
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
            object enumProxy = Invoker.PropertyGet(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(comObject, enumProxy, true);
            Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                COMObject returnClass = LateBindingApi.Core.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return returnClass;
            }
        }

        /// <summary>
        /// returns enumerator with com proxies
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetProxyEnumeratorAsMethod(COMObject comObject)
        {
            object enumProxy = Invoker.MethodReturn(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(comObject, enumProxy, true);
            Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                COMObject returnClass = LateBindingApi.Core.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return returnClass;
            }
        }

        /// <summary>
        /// returns enumerator with scalar variables
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetScalarEnumeratorAsProperty(COMObject comObject)
        {
            object enumProxy = Invoker.PropertyGet(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(comObject, enumProxy, true);
            Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object item = Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return item;
            }
        }

        /// <summary>
        /// returns enumerator with scalar variables
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static IEnumerator GetScalarEnumeratorAsMethod(COMObject comObject)
        {
            object enumProxy = Invoker.MethodReturn(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(comObject, enumProxy, true);
            Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object item = Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return item;
            }
        }

    }
}
