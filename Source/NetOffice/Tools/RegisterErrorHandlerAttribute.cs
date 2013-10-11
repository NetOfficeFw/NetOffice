using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Mark a static method as error handler for COMAddin methods. The static method need the following signature: public static void ErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
    /// Rethrow the exception(second argument) in the method body to the runtime system if you want signalize an error to the environment(typical not wanted)
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method)]
    public class RegisterErrorHandlerAttribute : System.Attribute
    {
    }

    /// <summary>
    /// Indicates in which method the error is occured
    /// </summary>
    public enum RegisterErrorMethodKind
    {
        /// <summary>
        /// the error is occured in the Register operation
        /// </summary>
        Register = 0,

        /// <summary>
        ///  the error is occured in the Unregister operation
        /// </summary>
        UnRegister = 1,
    }
}
