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
        #region Method

        /// <summary>
        /// perform method as latebind call 
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        public static void Method(COMObject comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        public static void Method(object comObject, string name)
        {
            Method(comObject, name, null);
        }

        /// <summary>
        /// perform method as latebind call 
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        public static void Method(COMObject comObject, string name, object[] paramsArray)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
        }

        /// <summary>
        /// perform method as latebind call 
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        public static void Method(object comObject, string name, object[] paramsArray)
        {
            comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, comObject, paramsArray, Settings.ThreadCulture);
        }

        /// <summary>
        /// perform method as latebind call 
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <param name="paramModifiers"></param>
        public static void Method(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static object MethodReturn(COMObject comObject, string name)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, null, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
        public static object MethodReturn(COMObject comObject, string name, object[] paramsArray)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform method as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <param name="paramModifiers"></param>
        /// <returns></returns>
        public static object MethodReturn(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
            return returnValue;
        }

        #endregion

        #region Property

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static object PropertyGet(object comObject, string name)
        {
            object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, null, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static object PropertyGet(COMObject comObject, string name)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, null, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
        public static object PropertyGet(object comObject, string name, object[] paramsArray)
        {
            object returnValue = comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, paramsArray, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
        public static object PropertyGet(COMObject comObject, string name, object[] paramsArray)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, Settings.ThreadCulture);
            return returnValue;
        }

        /// <summary>
        /// perform property get as latebind call with return value
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <param name="paramModifiers"></param>
        /// <returns></returns>
        public static object PropertyGet(COMObject comObject, string name, object[] paramsArray, ParameterModifier[] paramModifiers)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.UnderlyingObject, paramsArray, paramModifiers, Settings.ThreadCulture, null);
            return returnValue;
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <param name="value"></param>
        public static void PropertySet(COMObject comObject, string name, object[] paramsArray, object value)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = value;

            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray , Settings.ThreadCulture);
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="paramsArray"></param>
        /// <param name="value"></param>
        /// <param name="paramModifiers"></param>
        public static void PropertySet(COMObject comObject, string name, object[] paramsArray, object value, ParameterModifier[] paramModifiers)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = value;

            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, newParamsArray, paramModifiers, Settings.ThreadCulture, null);
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public static void PropertySet(COMObject comObject, string name, object value)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, Settings.ThreadCulture);
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="paramModifiers"></param>
        public static void PropertySet(COMObject comObject, string name, object value, ParameterModifier[] paramModifiers)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, new object[] { value }, paramModifiers, Settings.ThreadCulture, null);
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="paramModifiers"></param>
        public static void PropertySet(COMObject comObject, string name, object[] value, ParameterModifier[] paramModifiers)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, paramModifiers, Settings.ThreadCulture, null);
        }

        /// <summary>
        /// perform property set as latebind call
        /// </summary>
        /// <param name="comObject"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public static void PropertySet(COMObject comObject, string name, object[] value)
        {
            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.UnderlyingObject, value, Settings.ThreadCulture);
        }

        #endregion

        #region Parameters

        /// <summary>
        /// create parameter modifiers array
        /// </summary>
        /// <param name="isRef"></param>
        /// <returns></returns>
        public static ParameterModifier[] CreateParamModifiers(params bool[] isRef)
        {
            if (null != isRef)
            {
                ParameterModifier[] returnModifiers = new ParameterModifier[isRef.Length];
                for (int i = 0; i < isRef.Length; i++)
                {
                    ParameterModifier newModifiers = new ParameterModifier(1);
                    newModifiers[0] = isRef[i];
                    returnModifiers[i] = newModifiers;                    
                }

                return returnModifiers;
            }
            else
                return null;
        }

        /// <summary>
        /// replace null with Type.Missing, replace COMObject with COMObject.UnderlyingObject
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        public static object ValidateParam(object param)
        {
            if (null != param)
            {
                COMObject comObject = param as COMObject;
                if (null != comObject)
                        param = comObject.UnderlyingObject;

                if (param.GetType().IsEnum)
                    param = Convert.ToInt32(param);

                return param;
            }
            else
                return Type.Missing;
        }

        /// <summary>
        /// calls ValidateParam for every array item
        /// </summary>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
        public static object[] ValidateParamsArray(params object[] paramsArray)
        {
            if (null != paramsArray)
            {
                int parramArrayCount = paramsArray.Length;
                for (int i = 0; i<parramArrayCount; i++)
                    paramsArray[i] = ValidateParam(paramsArray[i]);
                return paramsArray;
            }
            else
                return null;
        }
        
        /// <summary>
        /// calls dipose in case if param is COMObject, calls Marshal.ReleaseComObject in case of param is a COM proxy
        /// </summary>
        /// <param name="param"></param>
        public static void ReleaseParam(object param)
        {
            if (null != param)
            {
                if (param is COMObject)
                {
                    COMObject comObject = param as COMObject;
                    comObject.Dispose();
                }
                else
                {
                    Type paramType = param.GetType();
                    if (true == paramType.IsCOMObject)
                        Marshal.ReleaseComObject(param);
                }
            }
        }

        /// <summary>
        /// calls ReleaseParam for every array item
        /// </summary>
        /// <param name="paramsArray"></param>
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
        /// creates an array with arrays
        /// </summary>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
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
        /// creates an array with arrays
        /// </summary>
        /// <param name="paramsModifier"></param>
        /// <param name="paramsArray"></param>
        /// <returns></returns>
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

        #region Type

        /// <summary>
        /// returns the Type for expression or null if param not set
        /// </summary>
        /// <param name="expression"></param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type GetObjectType(object expression)
        {
            if (null == expression)
                return null;
            else
                return expression.GetType();
        }

        #endregion
    }
}
