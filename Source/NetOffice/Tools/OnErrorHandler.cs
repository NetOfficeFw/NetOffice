using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Custom error handler
    /// </summary>
    /// <param name="methodKind">origin method where the error comes from</param>
    /// <param name="exception">occured exception</param>
    public delegate void OnErrorHandler(ErrorMethodKind methodKind, System.Exception exception);
}
