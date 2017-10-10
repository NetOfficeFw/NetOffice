using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a failed method invoke operation
    /// </summary>
    public class MethodCOMException : InvokerCOMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">inner exception</param>
        public MethodCOMException(string message, Exception innerException) : base(message, innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        public MethodCOMException(string message) : base(message, null)
        {

        }
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public MethodCOMException(Exception innerException) : base(null != innerException ? innerException.Message : "Failed to invoke property.", innerException)
        {

        }
    }
}
