﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Indicates a given com proxy/result doesn't implement the IDispatch COM Import Interface.
    /// The <see cref="IDispatch"/> interface is the key interface for late binding which NetOffice uses strictly.
    /// </summary>
    public class IDispatchNotImplementedException : NetOfficeException
    {
        /// <summary>
        /// Default Exception Message
        /// </summary>
        private static readonly string _defaultMessage = "Instance behind proxy doesn't implement IDispatch.";

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public IDispatchNotImplementedException() : base(_defaultMessage)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public IDispatchNotImplementedException(Exception innerException) : base(_defaultMessage, innerException)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public IDispatchNotImplementedException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public IDispatchNotImplementedException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
