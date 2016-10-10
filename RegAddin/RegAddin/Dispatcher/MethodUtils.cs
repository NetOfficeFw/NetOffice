using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

namespace RegAddin.Dispatcher
{
    internal static class MethodUtils
    {        
        private static string _comAddinBase = "NetOffice.Tools.COMAddinBase";

        internal static bool HasAttribute(MethodInfo info, string fullName)
        {
            object[] attributes = info.GetCustomAttributes(true);
            foreach (object item in attributes)
            {
                if (item.GetType().FullName == fullName)
                    return true;
            }
            return false;
        }

        internal static bool HasAttribute<T>(MethodInfo info) where T : System.Attribute
        {
            object[] attributes = info.GetCustomAttributes(typeof(T), false);
            return attributes.Length > 0;
        }

        internal static IEnumerable<MethodInfo> GetMethodsFromAddinBaseClass(Type item, BindingFlags flags)
        {           
            Type type = item;
            while (null != type)
            {              
                if (null != type.BaseType && type.BaseType.FullName == "System.Object" ||
                         type.BaseType.FullName == _comAddinBase)
                    break;
                type = type.BaseType;
            }

            return type.GetMethods(flags);
        }

        internal static IEnumerable<MethodInfo> GetMethods(Type item, BindingFlags flags)
        {
            List<MethodInfo> result = new List<MethodInfo>();
            Type type = item;
            while (null != type)
            {
                var methods = type.GetMethods(flags);
                foreach (var method in methods)
                {
                    result.Add(method);
                }
                if (null != type.BaseType && type.BaseType.FullName == "System.Object")
                    break;
                type = type.BaseType;
            }

            return result;
        }

        internal static object CallMethodWithArgumentsAndReturnValue(MethodInfo method, params object[] args)
        {
            try
            {
                return method.Invoke(null, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static bool CallMethodWithArguments(MethodInfo method, params object[] args)
        {
            try
            {
                method.Invoke(null, args);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static bool CallMethodWithArguments(MethodInfo method, Type addinType)
        {
            try
            {
                method.Invoke(null, new object[] { addinType });
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static bool CallMethodWithoutArguments(MethodInfo method)
        {
            try
            {
                method.Invoke(null, new object[0]);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
