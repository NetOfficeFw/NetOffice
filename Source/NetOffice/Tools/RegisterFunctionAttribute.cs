using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify the kind and condition of the registration method call.
    /// </summary>
    public enum RegisterMode
    {
        /// <summary>
        /// The base class <see cref="COMAddinBase"/> does not perform any registration operations
        /// and calls the specified registration method with the parameter <see cref="RegisterCall.Replace"/>.
        /// The specified registration method has to do all registration operations manually.
        /// </summary>
        Replace = 0,

        /// <summary>
        /// The method was called with parameter <see cref="RegisterCall.CallBefore"/> before the base class does any registration operations.
        /// </summary>
        CallBefore = 1,

        /// <summary>
        /// The method was called with parameter <see cref="RegisterCall.CallAfter"/> when the base class registration operations are done.
        /// </summary>
        CallAfter = 2,

        /// <summary>
        /// This means a combination of <see cref="RegisterCall.CallBefore"/> and <see cref="RegisterCall.CallAfter"/>.
        /// </summary>
        CallBeforeAndAfter = 3
    }

    /// <summary>
    /// Parameter for Register/Unregister Methods
    /// </summary>
    [System.Runtime.InteropServices.Guid("D8FAB9D7-10D1-4AA3-8DBA-D9CCA8C4CE9B")]
    public enum RegisterCall
    {
        /// <summary>
        /// The method was called without any register/unregister action from the base class. The specified register method has to do all register operations
        /// </summary>
        Replace = 0,

        /// <summary>
        /// The method is called before the base class perform all register operation
        /// </summary>
        CallBefore = 1,

        /// <summary>
        /// The method was called when the base class register operations are done
        /// </summary>
        CallAfter = 2,
    }

    /// <summary>
    /// Mark a static method as Register method. the method need the following signature public void Register(Type type, RegisterCall callType)
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method, AllowMultiple = false)]
    public class RegisterFunctionAttribute : System.Attribute
    {
        /// <summary>
        /// Register Call Condition
        /// </summary>
        public readonly RegisterMode Value;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="mode">register call condition</param>
        public RegisterFunctionAttribute(RegisterMode mode)
        {
            Value = mode;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        public RegisterFunctionAttribute()
        {

        }
    }
}
