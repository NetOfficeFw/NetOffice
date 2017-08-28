using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a Clone operation failed
    /// </summary>
    public class CloneException : NetOfficeException
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public CloneException(Exception innerException) : base("Failed to clone instance.", innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public CloneException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public CloneException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
