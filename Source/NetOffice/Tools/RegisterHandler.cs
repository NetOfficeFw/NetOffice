using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// Handle COMAddin register process
    /// </summary>
    public static class RegisterHandler
    {
        private static string ExceptionMessage = "An error occured while calling register.";

        /// <summary>
        /// Do register process 
        /// </summary>
        /// <param name="type">addin type</param>
        /// <param name="addinOfficeRegistryKey">office application registry path</param>
        /// <param name="scope">the current installation scope</param>
        /// <param name="keyState">the office registry key need to create</param>
        public static void Proceed(Type type, string[] addinOfficeRegistryKey, InstallScope scope, OfficeRegisterKeyState keyState)
        {
            try
            {
                GuidAttribute guid = AttributeReflector.GetGuidAttribute(type);
                ProgIdAttribute progId = AttributeReflector.GetProgIDAttribute(type);
                RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                COMAddinAttribute addin = AttributeReflector.GetCOMAddinAttribute(type);
                CodebaseAttribute codebase = AttributeReflector.GetCodebaseAttribute(type);
                LockbackAttribute lockBack = AttributeReflector.GetLockbackAttribute(type);
                ProgrammableAttribute programmable = AttributeReflector.GetProgrammableAttribute(type);
                bool isSystem = location.IsMachineTarget(scope);

                MethodInfo registerMethod = null;
                RegisterFunctionAttribute registerAttribute = null;
                bool registerMethodPresent = AttributeReflector.GetRegisterAttribute(type, ref registerMethod, ref registerAttribute);
                if (null != registerAttribute && true == registerMethodPresent &&
                    registerAttribute.Value == RegisterMode.CallBefore || registerAttribute.Value == RegisterMode.CallBeforeAndAfter)
                {
                    if (!CallDerivedRegisterMethod(registerMethod, type, registerAttribute.Value == RegisterMode.Replace ? RegisterCall.Replace : RegisterCall.CallBefore, scope, keyState))
                        if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, new NetOfficeException(ExceptionMessage)))
                            return;
                    if (registerAttribute.Value == RegisterMode.Replace)
                        return;
                }

                if (null != programmable)
                {
                    ProgrammableAttribute.CreateKeys(type.GUID, isSystem);
                }

                if (null != codebase && codebase.Value)
                {
                    Assembly thisAssembly = Assembly.GetAssembly(type);
                    string assemblyVersion = thisAssembly.GetName().Version.ToString();
                    CodebaseAttribute.CreateValue(type.GUID, location.IsMachineTarget(scope), assemblyVersion, thisAssembly.CodeBase);
                }

                if (null != lockBack)
                {                   
                    if (!LockbackAttribute.CreateKey(isSystem))
                        NetOffice.DebugConsole.Default.WriteLine("Unable to create lockback bypass.");
                }

                if (keyState == OfficeRegisterKeyState.NeedToCreate)
                {
                    foreach (string item in addinOfficeRegistryKey)
                    {                       
                        RegistryLocationAttribute.CreateApplicationKey(location.IsMachineTarget(scope), item, progId.Value,
                        addin.LoadBehavior, addin.Name, addin.Description, addin.CommandLineSafe);
                    }                    
                }

                if (null != registerAttribute && true == registerMethodPresent &&
                    registerAttribute.Value == RegisterMode.CallAfter || registerAttribute.Value == RegisterMode.CallBeforeAndAfter)
                {
                    if(!CallDerivedRegisterMethod(registerMethod, type, RegisterCall.CallAfter, scope, keyState))
                        RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, new NetOfficeException(ExceptionMessage));
                }
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception);
            }
        }

        /// <summary>
        /// Derived Register Call Helper
        /// </summary>
        /// <param name="registerMethod">the method to call</param>
        /// <param name="type">type for derived class</param>
        /// <param name="callType">kind of call, defined in Register attribute</param>
        /// <param name="scope">current register scope</param>
        /// <param name="keyState">office reg key state</param>
        /// <returns>true if no exception occurs</returns>
        private static bool CallDerivedRegisterMethod(MethodInfo registerMethod, Type type, 
            RegisterCall callType, InstallScope scope, OfficeRegisterKeyState keyState)
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
                        if(arguments[0].ParameterType.GUID == typeof(InstallScope).GUID)
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
                return true;
            }          
        }        
    }
}
