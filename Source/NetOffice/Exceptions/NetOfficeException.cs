﻿using System;

namespace NetOffice
{
    /// <summary>
    /// Indicates an exception occured in NetOffice.dll, not in corresponding NetOffice dependency assembly
    /// </summary>
    public class NetOfficeException : Exception
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        public NetOfficeException(string message) : base(message)
        {
            AppDomainId = AppDomain.CurrentDomain.Id;
            AppDomainFriendlyName = AppDomain.CurrentDomain.FriendlyName;
            AppDomainIsDefault = AppDomain.CurrentDomain.IsDefaultAppDomain();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerException">inner exception</param>
        public NetOfficeException( Exception innerException) : base(null != innerException? innerException.Message : "<NoErrorMessage>", innerException)
        {
            AppDomainId = AppDomain.CurrentDomain.Id;
            AppDomainFriendlyName = AppDomain.CurrentDomain.FriendlyName;
            AppDomainIsDefault = AppDomain.CurrentDomain.IsDefaultAppDomain();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">given exception info</param>
        /// <param name="innerException">inner exception</param>
        public NetOfficeException(string message, Exception innerException) : base(message, innerException)
        {
            AppDomainId = AppDomain.CurrentDomain.Id;
            AppDomainFriendlyName = AppDomain.CurrentDomain.FriendlyName;
            AppDomainIsDefault = AppDomain.CurrentDomain.IsDefaultAppDomain();
        }
        
        /// <summary>
        /// Current app domain is default app domain
        /// </summary>
        public bool AppDomainIsDefault { get; private set; }

        /// <summary>
        /// Id from current app domain
        /// </summary>
        public int AppDomainId { get; private set; }

        /// <summary>
        /// Friendly name from current app domain
        /// </summary>
        public string AppDomainFriendlyName { get; private set; }
    }
}