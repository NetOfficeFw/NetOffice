using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a child relation operation in Netoffice COM proxy management has been failed to complete
    /// </summary>
    public class COMChildRelationException : NetOfficeCOMException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">the exception that is the cause of the current exception</param>
        public COMChildRelationException(Exception innerException) : base(innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">the exception that is the cause of the current exception</param>
        public COMChildRelationException(string message, Exception innerException) : base(message, innerException)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        public COMChildRelationException(string message) : base(message)
        {

        }
    }
}