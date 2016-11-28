using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Handle register errors
    /// </summary>
    public static class RegisterErrorHandler
    {
        /// <summary>
        /// Checks for a static method, signed with the ErrorHandlerAttribute and call them if its available
        /// </summary>
        /// <param name="type">type information for the class wtih static method </param>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        /// <returns>true if error is handled by derived method an we can proceed</returns>
        public static bool RaiseStaticErrorHandlerMethod(Type type, RegisterErrorMethodKind methodKind, System.Exception exception)
        {
            MethodInfo errorMethod = AttributeReflector.GetRegisterErrorMethod(type);
            if (null != errorMethod)
            {
                try
                {
                    object result = null;
                    ParameterInfo[] arguments = errorMethod.GetParameters();
                    int argumentsCount = arguments.Length;
                    switch (argumentsCount)
                    {
                        case 0:
                            result = errorMethod.Invoke(null, new object[0]);
                            break;
                        case 1:
                            if(arguments[0].ParameterType.GUID == typeof(RegisterErrorMethodKind).GUID)
                                result = errorMethod.Invoke(null, new object[] { methodKind });
                            else
                                result = errorMethod.Invoke(null, new object[] { exception });
                            break;
                        case 2:
                            result = errorMethod.Invoke(null, new object[] { methodKind, exception });
                            break;
                        case 3:
                            result = errorMethod.Invoke(null, new object[] { type, methodKind, exception });
                            break;
                        default:
                            break;
                    }

                    if (result is bool)
                        return (bool)result;
                }
                catch (Exception throwedException)
                {
                    Console.WriteLine("Unable to call addin register error method. {0}", throwedException.Message);
                }
            }
            return false;
        }
    }
}
