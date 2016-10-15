using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Mark a static method as error handler for COMAddin methods. The static method need the following signature: public static void ErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
    /// Rethrow the exception(second argument) in the method body to the runtime system if you want signalize an error to the environment(typical not wanted)
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method, AllowMultiple = false)]
    public class RegisterErrorHandlerAttribute : System.Attribute
    {
    }

    /// <summary>
    /// Indicates in which method the error is occured
    /// </summary>
    [System.Runtime.InteropServices.Guid("F9A44508-4DC1-4E30-8195-0AFED88288E5")]
    public enum RegisterErrorMethodKind
    {
        /// <summary>
        /// The error is occured in the Register operation
        /// </summary>
        Register = 0,

        /// <summary>
        ///  The error is occured in the Unregister operation
        /// </summary>
        UnRegister = 1,

        /// <summary>
        /// The error is occured in the Register export operation
        /// </summary>
        Export = 2
    }
}
