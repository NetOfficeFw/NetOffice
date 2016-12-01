using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// NetOffice wrap all thrown exceptions from Office applications in a COMException. This enum can be used to set the exception message
    /// </summary>
    public enum ExceptionMessageHandling
    {
        ///// <summary>
        ///// The standard message from NetOffice.Settings.Exception
        ///// </summary>
        //Default = 0,

        /// <summary>
        /// The message from the source exception
        /// </summary>
        CopyInnerExceptionMessageToTopLevelException = 1,

        /// <summary>
        /// All inner exception messages as a summary
        /// </summary>
        CopyAllInnerExceptionMessagesToTopLevelException = 2
    }
}
