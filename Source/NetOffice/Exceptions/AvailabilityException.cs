using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates an availability operation failed
    /// </summary>
    public class AvailabilityException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public AvailabilityException(Exception innerException) : base("Failed to complete availability operation.", innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public AvailabilityException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public AvailabilityException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
