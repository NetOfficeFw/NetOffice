using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Represents an IDisposable state info instance. Dispose call do nothing if instance is already disposed
    /// </summary>
    public interface IDisposableState : IDisposable
    {
        /// <summary>
        /// Returns information the instance is already disposed
        /// </summary>
        bool IsDisposed { get; }
    }
}
