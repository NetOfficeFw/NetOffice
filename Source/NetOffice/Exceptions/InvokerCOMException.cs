﻿using System;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a failed invoke operation.
    /// </summary>
    public abstract class InvokerCOMException : NetOfficeCOMException
    {
        /// <summary>
        /// Creates an instance of the InvokerCOMException class.
        /// </summary>
        /// <param name="message">the message that indicates the reason for the exception</param>
        /// <param name="innerException">inner exception</param>
        protected InvokerCOMException(string message, Exception innerException) : base(message, innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the InvokerCOMException class.
        /// </summary>
        /// <param name="innerException">inner exception</param>
        protected InvokerCOMException(Exception innerException) : base(innerException?.Message ?? "Failed to invoke property.", innerException)
        {
        }
    }
}
