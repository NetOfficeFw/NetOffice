using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using System.Text;
using NetOffice.Exceptions;
using NetOffice.Availity;

namespace NetOffice
{
    /// <summary>
    /// Invokes ICOMObject instances
    /// </summary>
    public class Invoker
    {
        #region Fields

        /// <summary>
        /// lock field to perform thread safe operations
        /// </summary>
        private static object _lockInstance = new object();

        /// <summary>
        /// Shared default invoker
        /// </summary>
        private static Invoker _default;
         
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

        #region Method Invokes
        
        /// <summary>
        /// Perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        public void Method(ICOMObject comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// Perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        public void Method(object comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// Perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void Method(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);
              
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {             
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        public void Method(ICOMObject comObject, string name, object[] paramsArray, object value)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = ValidateParam(value);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, newParamsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void MethodWithoutSafeMode(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void Method(object comObject, string name, object[] paramsArray)
        {
            try
            {
                object target = null;
                Type type = null;

                ICOMObject wrapperInstance = comObject as ICOMObject;
                if (null != wrapperInstance)
                    ValidateComObjectIsAlive(wrapperInstance);

                if (null != wrapperInstance)
                {
                    target = wrapperInstance.UnderlyingObject;
                    type = wrapperInstance.UnderlyingType;
                }
                else
                {
                    target = comObject;
                    type = comObject.GetType();
                }

                bool measureStarted = false;
                if(null != wrapperInstance)
                    measureStarted = Settings.PerformanceTrace.StartMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                type.InvokeMember(name, BindingFlags.InvokeMethod, null, target, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters and parameter modifiers to use ref parameter(s)
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        public void Method(ICOMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <returns>any return value</returns>
        public object MethodReturn(ICOMObject comObject, string name)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object MethodReturn(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object MethodReturnWithoutSafeMode(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
               
                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Name, comObject.InstanceType.Name, name);
                
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object MethodReturn(ICOMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
                
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        #endregion

        #region Method (BindingFlags.InvokeMethod) Invokes

        /// <summary>
        /// Perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        public void SingleMethod(ICOMObject comObject, string name)
        {
            SingleMethod(comObject, name, null);
        }

        /// <summary>
        /// Perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        public void SingleMethod(object comObject, string name)
        {
            SingleMethod(comObject, name, null);
        }

        /// <summary>
        /// Perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void SingleMethod(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);
                
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void SingleMethodWithoutSafeMode(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public void SingleMethod(object comObject, string name, object[] paramsArray)
        {
            try
            {
                object target = null;
                Type type = null;

                ICOMObject wrapperInstance = comObject as ICOMObject;
                if (null != wrapperInstance)
                    ValidateComObjectIsAlive(wrapperInstance);

                if (null != wrapperInstance)
                {
                    target = wrapperInstance.UnderlyingObject;
                    type = wrapperInstance.UnderlyingType;
                }
                else
                {
                    target = comObject;
                    type = comObject.GetType();
                }

                bool measureStarted = false;
                if (null != wrapperInstance)
                    measureStarted = Settings.PerformanceTrace.StartMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                type.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, target, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with parameters and parameter modifiers to use ref parameter(s)
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        public void SingleMethod(ICOMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.Default.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Method);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(ICOMObject comObject, string name)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);
                
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object SingleMethodReturnWithoutSafeMode(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object SingleMethodReturn(ICOMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Method)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.Function);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);
                
                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new MethodCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.Method, paramsArray),
                    throwedException);
            }
        }

        #endregion

        #region Property Invokes

        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of property</param>
        /// <returns>any return value</returns>
        public object PropertyGet(object comObject, string name)
        {
            try
            {
                object target = null;
                Type type = null;

                ICOMObject wrapperInstance = comObject as ICOMObject;
                if (null != wrapperInstance)
                    ValidateComObjectIsAlive(wrapperInstance);

                if (null != wrapperInstance)
                {
                    target = wrapperInstance.UnderlyingObject;
                    type = wrapperInstance.UnderlyingType;
                }
                else
                {
                    target = comObject;
                    type = comObject.GetType();
                }

                bool measureStarted = false;
                if (null != wrapperInstance)
                    measureStarted = Settings.PerformanceTrace.StartMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = type.InvokeMember(name, BindingFlags.GetProperty, null, target, null, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <returns>any return value</returns>
        public object PropertyGet(ICOMObject comObject, string name)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);
             
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object PropertyGet(object comObject, string name, object[] paramsArray)
        {
            try
            {
                object target = null;
                Type type = null;

                ICOMObject wrapperInstance = comObject as ICOMObject;
                if (null != wrapperInstance)
                    ValidateComObjectIsAlive(wrapperInstance);

                if (null != wrapperInstance)
                {
                    target = wrapperInstance.UnderlyingObject;
                    type = wrapperInstance.UnderlyingType;
                }
                else
                {
                    target = comObject;
                    type = comObject.GetType();
                }

                bool measureStarted = false;
                if(null != wrapperInstance)
                    measureStarted = Settings.PerformanceTrace.StartMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = type.InvokeMember(name, BindingFlags.GetProperty, null, target, paramsArray, Settings.Default.ThreadCulture);
             
                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(wrapperInstance.InstanceType.Namespace, wrapperInstance.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet),
                    throwedException);
            }
        }
        
        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public object PropertyGet(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object PropertyGetWithoutSafeMode(ICOMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.Default.ThreadCulture);
                
                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet, paramsArray), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        /// <returns>any return value</returns>
        public object PropertyGet(ICOMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertyGet);

                object returnValue = comObject.UnderlyingType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.Default.ThreadCulture, null);
              
                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name);

                return returnValue;
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertyGetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertyGet, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        public void PropertySet(ICOMObject comObject, string name, object[] paramsArray, object value)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = ValidateParam(value);
              
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet, paramsArray),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>    
        public void PropertySet(ICOMObject comObject, string name, object[] paramsArray, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = value;
              
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet, paramsArray), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        public void PropertySet(ICOMObject comObject, string name, object value)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);
             
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, Settings.Default.ThreadCulture);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet),
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public void PropertySet(ICOMObject comObject, string name, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);
              
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value array to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public void PropertySet(ICOMObject comObject, string name, object[] value, ParameterModifier[] paramModifiers)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);
                
                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);

                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, paramModifiers, Settings.Default.ThreadCulture, null);

                if (measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet), 
                    throwedException);
            }
        }

        /// <summary>
        /// Perform property set as latebind call
        /// </summary>
        /// <param name="comObject">comobject instance</param>
        /// <param name="name">name of the property</param>
        /// <param name="value">new value of the property</param>
        public void PropertySet(ICOMObject comObject, string name, object[] value)
        {
            try
            {
                ValidateComObjectIsAlive(comObject);

                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportedEntityType.Property)))
                    throw new EntityNotSupportedException(name);

                bool measureStarted = Settings.PerformanceTrace.StartMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, PerformanceTrace.CallType.PropertySet);
                
                comObject.UnderlyingType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, Settings.Default.ThreadCulture);

                if(measureStarted)
                    Settings.PerformanceTrace.StopMeasureTime(comObject.InstanceType.Namespace, comObject.InstanceType.Name, name, value);
            }
            catch (Exception throwedException)
            {
                Console.WriteException(throwedException);
                throw new PropertySetCOMException(
                    ExceptionMessageBuilder.GetExceptionMessage(throwedException, comObject, name, CallType.PropertySet), 
                    throwedException);
            }
        }

        #endregion
        
        #region Parameters

        /// <summary>
        /// Create parameter modifiers array
        /// </summary>
        /// <param name="isRef">parameter is given as ref(ByRef in Visual Basic)</param>
        /// <returns>ParameterModifier array</returns>
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
        /// Replace null with Type.Missing, replace COMObject with COMObject.UnderlyingObject
        /// </summary>
        /// <param name="param">value to check</param>
        /// <returns>validated value</returns>
        public static object ValidateParam(object param)
        {
            if (null != param)
            {
                ICOMObject comObject = param as ICOMObject;
                if (!Object.ReferenceEquals(comObject, null))
                    param = comObject.UnderlyingObject;
                else if (param is Enum)
                    param = Convert.ToInt32(param);

                return param;
            }
            else
                return Type.Missing;
        }

        /// <summary>
        /// Calls ValidateParam for every array item
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
        /// Calls dipose in case if param is COMObject, calls Marshal.ReleaseComObject in case of param is a COM proxy
        /// </summary>
        public static void ReleaseParam(object param)
        {
            try
            {
                if (null != param)
                {
                    ICOMObject comObject = param as ICOMObject;
                    if (null != comObject)
                        comObject.Dispose();
                    else if (param is MarshalByRefObject)
                        Marshal.ReleaseComObject(param);
                }
            }
            catch (Exception throwedException)
            {
                DebugConsole.Default.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(
                    ExceptionMessageBuilder.GetDefaultExceptionMessage(),
                    throwedException);
            }
        }

        /// <summary>
        /// Calls ReleaseParam for every array item
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
        /// Copy the param array or returns null if paramsArray not set
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
        /// Copy the param array or returns null if paramsArray not set
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

        #region Methods

        /// <summary>
        /// Throws an ObjectDisposedException if given ICOMObject instance is already disposed.
        /// </summary>
        /// <param name="comObject">ICOMObject instance as any</param>
        private static void ValidateComObjectIsAlive(ICOMObject comObject)
        {
            if (comObject.IsDisposed || comObject.IsCurrentlyDisposing)
                throw new ObjectDisposedException("comObject");
        }

        #endregion
    }
}
