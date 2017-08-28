using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a failed property set invoke operation
    /// </summary>
    public class PropertySetCOMException : InvokerCOMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">inner exception</param>
        public PropertySetCOMException(string message, Exception innerException) : base(message, innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public PropertySetCOMException(Exception innerException) : base(null != innerException ? innerException.Message : "Failed to invoke property.", innerException)
        {

        }
    }
}
