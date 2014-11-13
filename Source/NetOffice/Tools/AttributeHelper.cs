using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;

namespace NetOffice.Tools
{
    /// <summary>
    /// Provides Attribute Helper functions
    /// </summary>
    [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
    public static class AttributeHelper
    {
        /// <summary>
        /// Anyalyze first parameter and returns the register error method delegate if exists
        /// </summary>
        /// <param name="type">Type of target addin</param>
        /// <returns>delegate or null</returns>
        public static MethodInfo GetRegisterErrorMethod(Type type)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public))
            {
                object[] array = item.GetCustomAttributes(typeof(RegisterErrorHandlerAttribute), false);
                if (array.Length == 1)
                {
                    ParameterInfo[] paramInfo = item.GetParameters();
                    if (paramInfo.Length == 2 && paramInfo[0].ParameterType == typeof(RegisterErrorMethodKind) && paramInfo[1].ParameterType == typeof(Exception))
                        return item;
                }
            }

            return null;
        }

        /// <summary>
        /// Looks for a method with the RegisterErrorHandlerFunctionAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the method</param>
        /// <param name="method">the method when its found</param>
        /// <param name="attribute">the attribute when its found</param>
        /// <returns>true when the method was found</returns>
        public static bool GetRegisterErrorAttribute(Type type, ref MethodInfo method, ref RegisterErrorHandlerAttribute attribute)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public))
            {
                object[] array = item.GetCustomAttributes(typeof(RegisterErrorHandlerAttribute), false);
                if (array.Length == 1)
                {
                    method = item;
                    attribute = array[0] as RegisterErrorHandlerAttribute;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        ///  Looks for a method with the RegisterFunctionAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the method</param>
        /// <param name="method">the method when its found</param>
        /// <param name="attribute">the attribute when its found</param>
        /// <returns>true when the method was found</returns>
        public static bool GetRegisterAttribute(Type type, ref MethodInfo method, ref RegisterFunctionAttribute attribute)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public))
            {
                object[] array = item.GetCustomAttributes(typeof(RegisterFunctionAttribute), false);
                if (array.Length == 1)
                {
                    method = item;
                    attribute = array[0] as RegisterFunctionAttribute;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Looks for a method with the UnRegisterFunctionAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the method</param>
        /// <param name="method">the method when its found</param>
        /// <param name="attribute">the attribute when its found</param>
        /// <returns>true when the method was found</returns>
        public static bool GetUnRegisterAttribute(Type type, ref MethodInfo method, ref UnRegisterFunctionAttribute attribute)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public))
            {
                object[] array = item.GetCustomAttributes(typeof(UnRegisterFunctionAttribute), false);
                if (array.Length == 1)
                {
                    method = item;
                    attribute = array[0] as UnRegisterFunctionAttribute;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Looks for the CustomUIAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>CustomUIAttribute or null</returns>
        public static CustomUIAttribute GetRibbonAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(CustomUIAttribute), false);
            if (array.Length == 0)
                return null;
            return array[0] as CustomUIAttribute;
        }

        /// <summary>
        /// Looks for the CustomPaneAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>CustomPaneAttribute or null</returns>
        public static CustomPaneAttribute GetCustomPaneAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(CustomPaneAttribute), false);
            if (array.Length == 0)
                return null;
            return array[0] as CustomPaneAttribute;
        }

        /// <summary>
        /// Looks for the GuidAttribute. Throws an exception if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>GuidAttribute</returns>
        public static GuidAttribute GetGuidAttribute(Type type)
        {

            object[] array = type.GetCustomAttributes(typeof(GuidAttribute), false);
            if (array.Length == 0)
                throw new ArgumentNullException("GuidAttribute is missing");
            return array[0] as GuidAttribute;
        }

        /// <summary>
        /// Looks for the ProgIdAttribute. Throws an exception if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>ProgIdAttribute</returns>
        public static ProgIdAttribute GetProgIDAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(ProgIdAttribute), false);
            if (array.Length == 0)
                throw new ArgumentNullException("ProgIdAttribute is missing");
            return array[0] as ProgIdAttribute;
        }

        /// Looks for the TweakAttribute.
        /// <summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// </summary>
        /// <returns>TweakAttribute</returns>
        public static TweakAttribute GetTweakAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(TweakAttribute), false);
            if (array.Length == 0)
                return new TweakAttribute(false);
            else
                return array[0] as TweakAttribute;
        }

        /// <summary>
        /// Looks for the RegistryLocationAttribute. Returns a default RegistryLocationAttribute(CurrentUser) if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>RegistryLocationAttribute</returns>
        public static RegistryLocationAttribute GetRegistryLocationAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(RegistryLocationAttribute), false);
            if (array.Length == 0)
                return new RegistryLocationAttribute(RegistrySaveLocation.CurrentUser);
            else
                return array[0] as RegistryLocationAttribute;
        }

        /// <summary>
        /// Looks for the COMAddinAttribute.
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>COMAddinAttribute</returns>
        public static COMAddinAttribute GetCOMAddinAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(COMAddinAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as COMAddinAttribute;
        }
    }
}