using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// invoke helper functions
    /// </summary>
    public static class Invoker
    {
        #region Fields

        private static readonly string _exceptionMessage = "See inner exception(s) for details.";

        #endregion

        #region Method

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        public static void Method(COMObject comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call without parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        public static void Method(object comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public static void Method(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if( (Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name,SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static void MethodWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters 
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        public static void Method(object comObject, string name, object[] paramsArray)
        {
            try
            { 
                comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject, paramsArray, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with parameters and parameter modifiers to use ref parameter(s)
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <param name="paramModifiers">ararry with modifiers correspond paramsArray</param>
        public static void Method(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <returns>any return value</returns>
        public static object MethodReturn(COMObject comObject, string name)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of method</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public static object MethodReturn(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static object MethodReturnWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static object MethodReturn(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Method)))
                    throw new EntityNotSupportedException(string.Format("Method {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static object PropertyGet(object comObject, string name)
        {
            try
            {
                object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, null, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <returns>any return value</returns>
        public static object PropertyGet(COMObject comObject, string name)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target proxy</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public static object PropertyGet(object comObject, string name, object[] paramsArray)
        {
            try
            {
                object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, paramsArray, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        
        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param>
        /// <returns>any return value</returns>
        public static object PropertyGet(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static object PropertyGetWithoutSafeMode(COMObject comObject, string name, object[] paramsArray)
        {
            try
            {
                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static object PropertyGet(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
                return returnValue;
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="paramsArray">array with parameters</param> 
        /// <param name="value">value to be set</param>
        public static void PropertySet(COMObject comObject, string name, object[] paramsArray, object value)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = value;

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
        public static void PropertySet(COMObject comObject, string name, object[] paramsArray, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                object[] newParamsArray = new object[paramsArray.Length + 1];
                for (int i = 0; i < paramsArray.Length; i++)
                    newParamsArray[i] = paramsArray[i];
                newParamsArray[newParamsArray.Length - 1] = value;

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, paramModifiers, Settings.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        public static void PropertySet(COMObject comObject, string name, object value)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public static void PropertySet(COMObject comObject, string name, object value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, paramModifiers, Settings.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject">target object</param>
        /// <param name="name">name of property</param>
        /// <param name="value">value array to be set</param>
        /// <param name="paramModifiers">array with modifiers correspond paramsArray</param>
        public static void PropertySet(COMObject comObject, string name, object[] value, ParameterModifier[] paramModifiers)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, paramModifiers, Settings.ThreadCulture, null);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
            }
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public static void PropertySet(COMObject comObject, string name, object[] value)
        {
            try
            {
                if ((Settings.EnableSafeMode) && (!comObject.EntityIsAvailable(name, SupportEntityType.Property)))
                    throw new EntityNotSupportedException(string.Format("Property {0} is not available.", name));

                comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, Settings.ThreadCulture);
            }
            catch (Exception throwedException)
            {
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
                if (null != comObject)
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
                DebugConsole.WriteException(throwedException);
                throw new System.Runtime.InteropServices.COMException(_exceptionMessage, throwedException);
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
    }
}
