using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Mark a static Method as Error Handler. The method need the following signature: public void bool ErrorHandler(Exception exception)
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method)]
    public class ErrorHandlerFunctionAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        public ErrorHandlerFunctionAttribute()
        { 
        }
    }
}
