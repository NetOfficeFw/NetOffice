using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Some helper methods (also for visual basic)
    /// </summary>
    public class Utils
    {
        #region Fields

        private static object _lockUtils = new object();
        private static Utils _default;

        #endregion

        #region Properties

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
                }
                return _default;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Checks value is null (Nothing in Visual Basic) 
        /// </summary>
        /// <param name="value">check value</param>
        /// <returns>true if null</returns>
        public static bool IsNothing(object value)
        {
            return null == value;
        }

        /// <summary>
        ///  Checks value is null (Nothing in Visual Basic) or Type.Missing
        /// </summary>
        /// <param name="value">check value</param>
        /// <returns>true if null or missing</returns>
        public static bool IsNullOrMissing(object value)
        {
            if (null == value || Type.Missing == value)
                return true;
            else
                return false;
        }
        
        /// <summary>
        /// Returns an enumerator with variant items - that means item(s) can be proxy or scalar
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetVariantEnumeratorAsProperty(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Variant Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    if (itemProxy is MarshalByRefObject)
                    { 
                        ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                        yield return returnClass;
                    }
                    else
                        yield return itemProxy;

                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator with com proxies
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetVariantEnumeratorAsMethod(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Variant Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    if (itemProxy is MarshalByRefObject)
                    {
                        ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                        yield return returnClass;
                    }
                    else
                        yield return itemProxy;

                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator with com proxies
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetProxyEnumeratorAsProperty(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Proxy Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                    yield return returnClass;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator with com proxies
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetProxyEnumeratorAsMethod(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Proxy Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy);
                    yield return returnClass;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator with scalar variables
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetScalarEnumeratorAsProperty(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Scalar Enumerator");
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
        /// Returns an enumerator with scalar variables
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetScalarEnumeratorAsMethod(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (_lockUtils)
            {
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Scalar Enumerator");
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

        #endregion
    }
}