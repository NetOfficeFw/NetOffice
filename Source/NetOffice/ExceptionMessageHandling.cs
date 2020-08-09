using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// NetOffice wraps all thrown exceptions from MS Office applications in <see cref="NetOfficeCOMException"/>.
    /// This enum can be used to change the exception message.
    /// </summary>
    public enum ExceptionMessageHandling
    {
        /// <summary>
        /// The standard message from <see cref="Settings.ExceptionDefaultMessage"/>.
        /// </summary>
        Default = 0,
        
        /// <summary>
        /// The message from the source exception.
        /// </summary>
        CopyInnerExceptionMessageToTopLevelException = 1,

        /// <summary>
        /// All inner exception messages as a summary.
        /// </summary>
        CopyAllInnerExceptionMessagesToTopLevelException = 2,

        /// <summary>
        /// The standard message from <see cref="Settings.ExceptionDiagnosticsMessage"/>.
        /// NetOffice will replace these placeholder strings in the diagnostics message (if they exist):
        /// {CallType}      - type of the call, such as method or property
        /// {CallInstance}  - friendly name of the COM object, using the <see cref="ICOMObjectProxy.InstanceFriendlyName"/>
        /// {Name}          - Name of the method or property
        /// {Args}          - given arguments
        /// </summary>
        Diagnostics = 3,
    }
}
