using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Signals the target method or property is not supported from the COM proxy in the current version
    /// </summary>
    public class EntityNotSupportedException : NetOfficeException 
    {        
        /// <summary>
        /// Creates an instance of the exception
        /// </summary>
        /// <param name="entityName">name of missing entity</param>
        public EntityNotSupportedException(string entityName) : base(String.Format("Not available:{0}.", entityName))
        {

        }
    }
}
