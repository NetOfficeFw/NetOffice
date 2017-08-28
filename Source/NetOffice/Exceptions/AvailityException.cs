using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates an availity operation failed
    /// </summary>
    public class AvailityException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public AvailityException(Exception innerException) : base("Failed to complete availity operation.", innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public AvailityException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public AvailityException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
