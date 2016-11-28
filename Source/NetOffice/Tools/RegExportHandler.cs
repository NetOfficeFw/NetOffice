using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Handle COMAddin register export process
    /// </summary>
    public class RegExportHandler
    {
        /// <summary>
        /// Do register export process per user
        /// </summary>
        /// <param name="type">addin type</param>
        /// <param name="addinOfficeRegistryKey">office application registry path</param>
        /// <param name="keyState">the office registry key need to create</param>
        public static RegExport ProceedUser(Type type, string[] addinOfficeRegistryKey, OfficeRegisterKeyState keyState)
        {
            return Proceed(type, addinOfficeRegistryKey, InstallScope.User, keyState);
        }

        /// <summary>
        /// Do register export process
        /// </summary>
        /// <param name="type">addin type</param>
        /// <param name="addinOfficeRegistryKey">office application registry path</param>
        /// <param name="scope">the current installation scope</param>
        /// <param name="keyState">the office registry key need to create</param>
        public static RegExport Proceed(Type type, string[] addinOfficeRegistryKey, InstallScope scope, OfficeRegisterKeyState keyState)
        {            
            try
            {
                object result = null;
                MethodInfo exportMethod = null;
                RegExportFunctionAttribute registerAttribute = null;
                bool registerMethodPresent = AttributeReflector.GetRegExportAttribute(type, ref exportMethod, ref registerAttribute);
                if (registerMethodPresent)
                {
                    ParameterInfo[] arguments = exportMethod.GetParameters();
                    int argumentCount = arguments.Length;
                    switch (argumentCount)
                    {
                        case 0:
                            result = exportMethod.Invoke(null, new object[0]);
                            break;
                        case 1:
                            result = exportMethod.Invoke(null, new object[] { scope });
                            break;
                        case 2:
                            result = exportMethod.Invoke(null, new object[] { scope, keyState});
                            break;
                        case 3:
                            exportMethod.Invoke(null, new object[] { type, scope, keyState });
                            break;
                        default:
                            break;
                    }

                    return result as RegExport;
                }
                else
                    return null;
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Export, exception);
                return null;
            }
        }
    }
}
