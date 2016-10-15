using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify the kind and condition of an register method call
    /// </summary>
    public enum RegisterMode
    {
        /// <summary>
        /// the base class COMAddin doesnt perform any register operations and call the specified register method with the parameter RegisterCall.Replace. The specified register method has to do all register operations
        /// </summary>
        Replace = 0,

        /// <summary>
        /// the method was called with parameter RegisterCall.CallBefore before the base class do any register operations. 
        /// </summary>
        CallBefore = 1,

        /// <summary>
        /// the method was called with parameter RegisterCall.CallAfter when the base register operations is done. 
        /// </summary>
        CallAfter = 2,

        /// <summary>
        /// this means a combination of CallBefore and CallAfter
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
        /// the method was called without any register/unregister action from the base class. The specified register method has to do all register operations
        /// </summary>
        Replace = 0,

        /// <summary>
        /// the method is called before the base class perform all register operation
        /// </summary>
        CallBefore = 1,

        /// <summary>
        /// the method was called when the base class register operations are done
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
    }
}
