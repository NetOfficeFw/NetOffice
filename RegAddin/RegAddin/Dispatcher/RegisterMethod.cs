using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RegAddin.Dispatcher
{
    internal class RegisterMethod
    {
        private string _NetOfficeRegisterName = "NetOffice.Tools.ComRegisterCallAttribute";

        internal bool Call(Type addinType, int installScope, int keyState)
        {
            bool result = true;

            IEnumerable<MethodInfo> methods = MethodUtils.GetMethods(addinType, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            MethodInfo method = TryGetNetOfficeRegisterMethod(methods);
            if (null != method)
            {
                ParameterInfo[] arguments = method.GetParameters();
                if (arguments.Length == 3)
                {
                    if (!MethodUtils.CallMethodWithArguments(method, addinType, installScope, keyState))
                        result = false;
                }
            }
            else
            {
                method = TryGetComRegisterMethod(methods);
                if (null != method)
                    return result;

                ParameterInfo[] arguments = method.GetParameters();
                if (arguments.Length == 0)
                {
                    if (!MethodUtils.CallMethodWithoutArguments(method))
                        result = false;
                }
                else if (arguments.Length == 1 && arguments[0].ParameterType == typeof(Type))
                {
                    if (!MethodUtils.CallMethodWithArguments(method, addinType))
                        result = false;
                }
            }
            
            return result;
        }
        
        private MethodInfo TryGetNetOfficeRegisterMethod(IEnumerable<MethodInfo> methods)
        {
            foreach (MethodInfo item in methods)
            {
                if (MethodUtils.HasAttribute(item, _NetOfficeRegisterName))
                    return item;
            }
            return null;
        }

        private MethodInfo TryGetComRegisterMethod(IEnumerable<MethodInfo> methods)
        {
            foreach (MethodInfo item in methods)
            {
                if (MethodUtils.HasAttribute<ComRegisterFunctionAttribute>(item))
                    return item;
            }
            return null;
        }
    }
}
