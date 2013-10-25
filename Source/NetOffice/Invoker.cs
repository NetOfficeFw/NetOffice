using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Invoke helper functions
    /// </summary>
    public class Invoker
    {
        #region Fields

        /// <summary>
        /// lock field to perform thread safe operations
        /// </summary>
        private static object _lockInstance = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parentFactory">parent factory</param>
        internal Invoker(Core parentFactory)
        {
            Parent = parentFactory;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal Invoker()
        {
            IsDefault = true;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Shared Default Invoker
        /// </summary>
        public static Invoker Default
        {
            get
            {
                lock (_lockInstance)
                {
                    if (null == _default)
                        _default = new Invoker();
                    return _default;
                }
            }
        }
        private static Invoker _default;

        /// <summary>
        /// Returns info this invoker is the default instance
        /// </summary>
        public bool IsDefault { get; private set; }

        /// <summary>
        /// Parent Factory
        /// </summary>
        internal Core Parent { get; private set; }

        /// <summary>
        /// Associated DebugConsole
        /// </summary>
        internal DebugConsole Console
        {
            get
            {
                if (null != Parent)
                    return Parent.Console;
                else
                    return DebugConsole.Default;
            }
        }

        /// <summary>
        /// Associated Settings
        /// </summary>
        internal Settings Settings
        {
            get
            {
                if (null != Parent)
                    return Parent.Settings;
                else
                    return Settings.Default;
            }
        }

        #endregion

        #region Method

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        public void Method(COMObject comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        public void Method(object comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void Method(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void MethodWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void Method(object comObject, string name, object[] paramsArray)
        {
            try
            {
                if ((comObject as COMObject).IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, comObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters and parameter modifiers to use ref parameter(s)
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        public void Method(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <returns>any return value</returns>
        public object MethodReturn(COMObject comObject, string name)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object MethodReturn(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object MethodReturnWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object MethodReturn(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        #endregion

        #region Method (BindingFlags.InvokeMethod)

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        public void SingleMethod(COMObject comObject, string name)
        {
            SingleMethod(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        public void SingleMethod(object comObject, string name)
        {
            SingleMethod(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void SingleMethod(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void SingleMethodWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void SingleMethod(object comObject, string name, object[] paramsArray)
        {
            try
            {
                if ((comObject as COMObject).IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject, paramsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters and parameter modifiers to use ref parameter(s)
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        public void SingleMethod(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.Default.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(COMObject comObject, string name)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object SingleMethodReturnWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        #endregion

        #region Property

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of property</param>
        /// <returns>any return value</returns>
        public object PropertyGet(object comObject, string name)
        {
            try
            {
                if ((comObject as COMObject).IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, null, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <returns>any return value</returns>
        public object PropertyGet(COMObject comObject, string name)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object PropertyGet(object comObject, string name, object[] paramsArray)
        {
            try
            {
                if ((comObject as COMObject).IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object PropertyGet(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object PropertyGetWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }


        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object PropertyGet(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        public void PropertySet(COMObject comObject, string name, object[] paramsArray, object value)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = value;

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>    
        public void PropertySet(COMObject comObject, string name, object[] paramsArray, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = value;

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, paramModifiers, Settings.Default.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        public void PropertySet(COMObject comObject, string name, object value)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public void PropertySet(COMObject comObject, string name, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, paramModifiers, Settings.Default.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value array to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public void PropertySet(COMObject comObject, string name, object[] value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, paramModifiers, Settings.Default.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">comobject instance</param>
        /// <param name="name">name of the property</param>
        /// <param name="value">new value of the property</param>
        public void PropertySet(COMObject comObject, string name, object[] value)
        {
            try
            {
                if (comObject.IsDisposed)
                    throw new ObjectDisposedException("COMObject");

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, Settings.Default.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        #endregion

        #region Parameters

        /// <summary>
        /// create parameter modifiers array
        /// </summary>
        /// <param name="isRef">parameter is given as ref(ByRef in Visual Basic)</param>
        /// <returns></returns>
        public static ParameterModifier[] CreateParamModifiers(params bool[] isRef)
        {
            if (null != isRef)
            {
                ParameterModifier arrPmods = new ParameterModifier(isRef.Length);
                for (int i = 0; i < isRef.Length; i++)
                    arrPmods[i] = isRef[i];

                ParameterModifier[] returnModifiers = { arrPmods };
                return returnModifiers;
            }
            else
                return null;
        }

        /// <summary>
        /// replace null with Type.Missing, replace COMObject with COMObject.UnderlyingObject
        /// </summary>
        /// <param name="param">value to check</param>
        /// <returns>validated value</returns>
        public static object ValidateParam(object param)
        {
            if (null != param)
            {
                COMObject comObject = param as COMObject;

                if (!Object.ReferenceEquals(comObject, null))
                    param = comObject.UnderlyingObject;
                else if (param.GetType().IsEnum)
                    param = Convert.ToInt32(param);

                return param;
            }
            else
                return Type.Missing;
        }

        /// <summary>
        /// calls ValidateParam for every array item
        /// </summary>
        /// <param name="paramsArray">array with parameters</param>
        public static object[] ValidateParamsArray(params object[] paramsArray)
        {
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i < parramArrayCount; i++)
                    paramsArray[i] = ValidateParam(paramsArray[i]);
                return paramsArray;
            }
            else
                return null;
        }

        /// <summary>
        /// calls dipose in case if param is COMObject, calls Marshal.ReleaseComObject in case of param is a COM proxy
        /// </summary>
        public static void ReleaseParam(object param)
        {
            try
            {
                if (null != param)
                {
                    COMObject comObject = param as COMObject;
                    if (null != comObject)
                        comObject.Dispose();
                    else if (param is MarshalByRefObject)
                        Marshal.ReleaseComObject(param);
                }
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(GetExceptionMessage(throwedException), throwedException);
            }
        }

        /// <summary>
        /// calls ReleaseParam for every array item
        /// </summary>
        /// <param name="paramsArray">any value array</param>
        public static void ReleaseParamsArray(params object[] paramsArray)
        {
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i < parramArrayCount; i++)
                    ReleaseParam(paramsArray[i]);
            }
        }

        /// <summary>
        /// copy the param array or returns null if paramsArray not set
        /// </summary>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>array copy or null</returns>
        public static object[] CreateEventParamsArray(params object[] paramsArray)
        {
            object[] returnArray = null;
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i < parramArrayCount; i++)
                    returnArray[i] = paramsArray[i];
                return returnArray;
            }
            else
                return null;
        }

        /// <summary>
        /// copy the param array or returns null if paramsArray not set
        /// </summary>
        /// <param name="paramsModifier">ararry with modifiers correspond paramsArray</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>array copy or null</returns>
        public static object[] CreateEventParamsArray(bool[] paramsModifier, params object[] paramsArray)
        {
            object[] returnArray = null;
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i < parramArrayCount; i++)
                {
                    if (true == paramsModifier[i])
                        returnArray[i] = paramsArray[i];
                    else
                        returnArray.SetValue(paramsArray[i], i);
                }
                return returnArray;
            }
            else
                return null;
        }

        #endregion

        #region Privates

        private static string GetExceptionMessage(Exception throwedException)
        {
            switch (Settings.Default.UseExceptionMessage)
            {
                case ExceptionMessageHandling.Default:

                    return Settings.Default.ExceptionMessage;

                case ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException:

                    string message = string.Empty;
                    while (throwedException.InnerException != null)
                    {
                        message = throwedException.Message;
                        throwedException = throwedException.InnerException;
                    }
                    return message;

                case ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException:

                    string messageSummary = string.Empty;
                    while (throwedException.InnerException != null)
                    {
                        messageSummary += throwedException.Message + Environment.NewLine;
                        throwedException = throwedException.InnerException;
                    }
                    return messageSummary;

                default:
                    throw new NetOfficeException("ArgumentOutOfRange:Settings.CopyInnerExceptionMessage");
            }
        }

        #endregion
    }
}
