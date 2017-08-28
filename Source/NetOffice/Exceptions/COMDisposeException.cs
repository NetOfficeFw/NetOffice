using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a dispose operation has been failed to complete
    /// </summary>
    public class COMDisposeException : NetOfficeCOMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">the exception that is the cause of the current exception</param>
        public COMDisposeException(Exception innerException) : base(innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">the exception that is the cause of the current exception</param>
        public COMDisposeException(string message, Exception innerException) : base(message, innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        public COMDisposeException(string message) : base(message)
        {

        }
    }
}