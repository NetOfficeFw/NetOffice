using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Signals the target method or property is not supported from the COM proxy in the current version
    /// </summary>
    public class EntityNotSupportedException : NetOfficeException 
    {        
        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="message"></param>
        public EntityNotSupportedException(string message): base(message)
        { }
    }
}
