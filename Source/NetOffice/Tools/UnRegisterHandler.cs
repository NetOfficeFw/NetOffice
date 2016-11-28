using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// Handle COMAddin unregister process
    /// </summary>
    public static class UnRegisterHandler
    {
        private static string _exceptionMessage = "An error occured while calling unregister.";
        
        /// <summary>
        /// Do unregister process per user uninstallation
        /// </summary>
        /// <param name="type">addin type</param>
        /// <param name="addinOfficeRegistryKey">office application registry path</param>
        /// <param name="keyState">the office registry key need to delete</param>
        public static void ProceedUser(Type type, string[] addinOfficeRegistryKey, OfficeUnRegisterKeyState keyState)
        {
            Proceed(type, addinOfficeRegistryKey, InstallScope.User, keyState);
        }

        /// <summary>
        /// Do unregister process 
        /// </summary>
        /// <param name="type">addin type</param>
        /// <param name="addinOfficeRegistryKey">office application registry path</param>
        /// <param name="scope">the current installation scope</param>
        /// <param name="keyState">the office registry key need to delete</param>
        public static void Proceed(Type type, string[] addinOfficeRegistryKey, InstallScope scope, OfficeUnRegisterKeyState keyState)
        {
            try
            {                 
                MethodInfo registerMethod = null;
                UnRegisterFunctionAttribute registerAttribute = null;
                bool registerMethodPresent = AttributeReflector.GetUnRegisterAttribute(type, ref registerMethod, ref registerAttribute);

                if ((null != registerAttribute && true == registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallBefore || registerAttribute.Value == RegisterMode.CallBeforeAndAfter))
                {
                    if (!CallDerivedUnRegisterMethod(registerMethod, type, registerAttribute.Value == RegisterMode.Replace ? RegisterCall.Replace : RegisterCall.CallBefore, scope, keyState))
                        if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, new NetOfficeException(_exceptionMessage)))
                            return;
                    if (registerAttribute.Value == RegisterMode.Replace)
                        return;
                }
                  
                ProgIdAttribute progId = AttributeReflector.GetProgIDAttribute(type);
                RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                CodebaseAttribute codebase = AttributeReflector.GetCodebaseAttribute(type);
                ProgrammableAttribute programmable = AttributeReflector.GetProgrammableAttribute(type);
                bool isSystemComponent = location.IsMachineComponentTarget(scope);
                bool isSystemAddin = location.IsMachineAddinTarget(scope);

                if (null != programmable)
                {
                    if(!ProgrammableAttribute.DeleteKeys(type.GUID, isSystemComponent, false))
                        NetOffice.DebugConsole.Default.WriteLine("Failed to delete programmable.");
                }

                if (null != codebase && codebase.Value == true)
                {
                    Assembly thisAssembly = Assembly.GetAssembly(type);
                    string assemblyVersion = thisAssembly.GetName().Version.ToString();
                    if (!CodebaseAttribute.DeleteValue(type.GUID, isSystemComponent, assemblyVersion, false))
                        NetOffice.DebugConsole.Default.WriteLine("Failed to delete codebase.");
                }

                if (keyState == OfficeUnRegisterKeyState.NeedToDelete)
                {                    
                    foreach (string item in addinOfficeRegistryKey)
                    {
                        RegistryLocationAttribute.TryDeleteApplicationKey(isSystemAddin, item, progId.Value);
                    }
                }

                if ((null != registerAttribute && true == registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallAfter || registerAttribute.Value == RegisterMode.CallBeforeAndAfter))
                {
                    if (!CallDerivedUnRegisterMethod(registerMethod, type, RegisterCall.CallAfter, scope, keyState))
                        RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, new NetOfficeException(_exceptionMessage));
                }
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, exception);
            }
        }

        /// <summary>
        /// Derived UnRegister Call Helper
        /// </summary>
        /// <param name="registerMethod">the method to call</param>
        /// <param name="type">type for derived class</param>
        /// <param name="callType">kind of call, defined in Register attribute</param>
        /// <param name="scope">current register scope</param>
        /// <param name="keyState">office reg key state</param>
        /// <returns>true if no exception occurs, otherwise false</returns>
        private static bool CallDerivedUnRegisterMethod(MethodInfo registerMethod, Type type,
            RegisterCall callType, InstallScope scope, OfficeUnRegisterKeyState keyState)
        {
            try
            {
                ParameterInfo[] arguments = registerMethod.GetParameters();
                int argumentsCount = arguments.Length;
                switch (argumentsCount)
                {
                    case 0:
                        registerMethod.Invoke(null, new object[0]);
                        break;
                    case 1:
                        if (arguments[0].ParameterType.GUID == typeof(InstallScope).GUID)
                            registerMethod.Invoke(null, new object[] { scope });
                        else if (arguments[0].ParameterType.GUID == typeof(RegisterCall).GUID)
                            registerMethod.Invoke(null, new object[] { callType });
                        else
                            registerMethod.Invoke(null, new object[] { type });
                        break;
                    case 2:
                        registerMethod.Invoke(null, new object[] { type, callType });
                        break;
                    case 3:
                        registerMethod.Invoke(null, new object[] { type, callType, scope });
                        break;
                    case 4:
                        registerMethod.Invoke(null, new object[] { type, callType, scope, keyState });
                        break;
                    default:
                        break;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
     }
}
