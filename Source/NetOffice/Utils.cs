using System;
using System.Collections;

namespace NetOffice
{
    /// <summary>
    /// Some helper methods (also for visual basic)
    /// The main purpose is accessing COM enumerators.
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
        /// Creates a managed enumerator
        /// </summary>
        ///  <param name="parent">parent instance or null in com proxy management</param>
        /// <param name="comObject">ICOMObject instance to access the enumerator</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>managed enumerator</returns>
        public static ICOMObject GetComObjectEnumeratorAsProperty(ICOMObject parent, ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, parent, enumProxy, true, "Variant Enumerator");
                return enumerator;
            }
        }

        /// <summary>
        /// Creates a managed enumerator
        /// </summary>
        ///  <param name="parent">parent instance or null in com proxy management</param>
        /// <param name="comObject">ICOMObject instance to access the enumerator</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>managed enumerator</returns>
        public static ICOMObject GetComObjectEnumeratorAsMethod(ICOMObject parent, ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, parent, enumProxy, true, "Variant Enumerator");
                return enumerator;
            }
        }

        /// <summary>
        /// Fetch managed enumerator
        /// </summary>
        /// <param name="parent">parent instance or null in com proxy management</param>
        /// <param name="enumerator">enumerator to fetch</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerable FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator, bool allowDynamicObject)
        {
            lock (enumerator.SyncRoot)
            {
                enumerator.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)enumerator.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = enumerator.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    if (itemProxy is MarshalByRefObject)
                    {
                        ICOMObject returnClass = enumerator.Factory.CreateObjectFromComProxy(parent, itemProxy, allowDynamicObject);
                        yield return returnClass;
                    }
                    else
                        yield return itemProxy;

                    isMoveNextTrue = (bool)enumerator.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }
        
        /// <summary>
        /// Returns an enumerator with com proxies
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetDuckVariantEnumeratorAsProperty(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
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
                        ICOMObject returnClass = comObject.Factory.CreateDuckObjectFromComProxy(enumerator, itemProxy);
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
        public static IEnumerator GetDuckVariantEnumeratorAsMethod(ICOMObject comObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
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
                        ICOMObject returnClass = comObject.Factory.CreateDuckObjectFromComProxy(enumerator, itemProxy);
                        yield return returnClass;
                    }
                    else
                        yield return itemProxy;

                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }
        
        /// <summary>
        /// Returns an enumerator with variant items - that means item(s) can be proxy or scalar
        /// </summary>
        /// <param name="comObject">ICOMObject instance as any</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetVariantEnumeratorAsProperty(ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Variant Enumerator");
                enumerator.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)enumerator.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    if (itemProxy is MarshalByRefObject)
                    { 
                        ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy, allowDynamicObject);
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
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetVariantEnumeratorAsMethod(ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Variant Enumerator");
                enumerator.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    if (itemProxy is MarshalByRefObject)
                    {
                        ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy, allowDynamicObject);
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
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetProxyEnumeratorAsProperty(ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.PropertyGet(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Proxy Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy, allowDynamicObject);
                    yield return returnClass;
                    isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator with com proxies
        /// </summary>
        /// <param name="comObject">COMObject instance as any</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>IEnumerator instance</returns>
        public static IEnumerator GetProxyEnumeratorAsMethod(ICOMObject comObject, bool allowDynamicObject)
        {
            if (null == comObject)
                throw new ArgumentNullException("comObject");

            lock (comObject.SyncRoot)
            {
                if (!comObject.Settings.EnableDynamicObjects)
                    allowDynamicObject = false;
                comObject.Factory.CheckInitialize();
                object enumProxy = comObject.Factory.Invoker.MethodReturn(comObject, "_NewEnum");
                ICOMObject enumerator = new COMObject(comObject.Factory, comObject, enumProxy, true, "Proxy Enumerator");
                comObject.Factory.Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
                bool isMoveNextTrue = (bool)comObject.Factory.Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                while (true == isMoveNextTrue)
                {
                    object itemProxy = comObject.Factory.Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                    ICOMObject returnClass = comObject.Factory.CreateObjectFromComProxy(enumerator, itemProxy, allowDynamicObject);
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

            lock (comObject.SyncRoot)
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

            lock (comObject.SyncRoot)
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