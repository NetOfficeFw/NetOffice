using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;

namespace NetOffice.Tools
{
    /// <summary>
    /// Provides Attribute reflection utils
    /// </summary>
    [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
    public static class AttributeReflector
    {
        /// <summary>
        /// Anyalyze first parameter and returns the register error method delegate if exists
        /// </summary>
        /// <param name="type">Type of target addin</param>
        /// <returns>delegate or null</returns>
        public static MethodInfo GetRegisterErrorMethod(Type type)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
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
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
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
        ///  Looks for a static method with the RegExportFunctionAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the method</param>
        /// <param name="method">the method when its found</param>
        /// <param name="attribute">the attribute when its found</param>
        /// <returns>true when the method was found</returns>
        public static bool GetRegExportAttribute(Type type, ref MethodInfo method, ref RegExportFunctionAttribute attribute)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
            {
                object[] array = item.GetCustomAttributes(typeof(RegExportFunctionAttribute), false);
                if (array.Length == 1)
                {
                    method = item;
                    attribute = array[0] as RegExportFunctionAttribute;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        ///  Looks for a static method with the RegisterFunctionAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the method</param>
        /// <param name="method">the method when its found</param>
        /// <param name="attribute">the attribute when its found</param>
        /// <returns>true when the method was found</returns>
        public static bool GetRegisterAttribute(Type type, ref MethodInfo method, ref RegisterFunctionAttribute attribute)
        {
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
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
            foreach (MethodInfo item in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
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
        /// Looks the CustomPaneAttributes
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>CustomPaneAttribute[] instance</returns>
        public static CustomPaneAttribute[] GetCustomPaneAttributes(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(CustomPaneAttribute), false);
            if (array.Length > 0)
            {
                CustomPaneAttribute[] result = new CustomPaneAttribute[array.Length];
                for (int i = 0; i < array.Length; i++)
                    result[i] = array[i] as CustomPaneAttribute;
                return result;
            }
            else
            { 
                return new CustomPaneAttribute[0];
            }
        }

        /// <summary>
        /// Looks for the TimestampAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>TimestampAttribute or null</returns>
        public static TimestampAttribute GetTimestampAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(TimestampAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as TimestampAttribute;
        }

        /// <summary>
        /// Looks for the ProgrammableAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>ProgrammableAttribute or null</returns>
        public static ProgrammableAttribute GetProgrammableAttribute(Type type)
        {

            object[] array = type.GetCustomAttributes(typeof(ProgrammableAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as ProgrammableAttribute;
        }


        /// <summary>
        /// Looks for the LockbackAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>LockbackAttribute or null</returns>
        public static LockbackAttribute GetLockbackAttribute(Type type)
        {

            object[] array = type.GetCustomAttributes(typeof(LockbackAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as LockbackAttribute;
        }

        /// <summary>
        /// Looks for the CodebaseAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>CodebaseAttribute or default attribute</returns>
        public static CodebaseAttribute GetCodebaseAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(CodebaseAttribute), false);
            if (array.Length == 0)
                return new CodebaseAttribute(true);
            else
                return array[0] as CodebaseAttribute;
        }

        /// <summary>
        /// Looks for the ForceInitializeAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>ForceInitializeAttribute or null</returns>
        public static ForceInitializeAttribute GetForceInitializeAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(ForceInitializeAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as ForceInitializeAttribute;
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
                throw new ArgumentOutOfRangeException(nameof(type), "GuidAttribute is missing");
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
                throw new ArgumentOutOfRangeException(nameof(type), "ProgIdAttribute is missing");
            return array[0] as ProgIdAttribute;
        }

        /// <summary>
        /// Looks for the ProgIdAttribute. Throws an exception if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <param name="throwException">throw exception if not found</param>
        /// <returns>ProgIdAttribute</returns>
        public static ProgIdAttribute GetProgIDAttribute(Type type, bool throwException)
        {
            object[] array = type.GetCustomAttributes(typeof(ProgIdAttribute), false);
            if (array.Length == 0)
            {
                if (throwException)
                    throw new ArgumentOutOfRangeException(nameof(type), "ProgIdAttribute is missing");
                else
                    return null;
            }
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
        /// Looks for the RegistryLocationAttribute. Returns a default RegistryLocationAttribute(InstallScope) if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>RegistryLocationAttribute</returns>
        public static RegistryLocationAttribute GetRegistryLocationAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(RegistryLocationAttribute), false);
            if (array.Length == 0)
                return new RegistryLocationAttribute(RegistrySaveLocation.InstallScope);
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
                throw new ArgumentOutOfRangeException(nameof(type), "COMAddinAttribute is missing");
            else
                return array[0] as COMAddinAttribute;
        }

        /// <summary>
        /// Looks for the COMAddinAttribute.
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <param name="progID">addin progid</param>
        /// <returns>COMAddinAttribute</returns>
        public static COMAddinAttribute GetCOMAddinAttribute(Type type, string progID)
        {
            object[] array = type.GetCustomAttributes(typeof(COMAddinAttribute), false);
            if (array.Length == 0)
                return new COMAddinAttribute(progID, String.Empty, 3);
            else
                return array[0] as COMAddinAttribute;
        }
    }
}
