using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Represents various dispose possibilities
    /// </summary>
    public interface ICOMObjectDisposable : IDisposable
    {
        /// <summary>
        /// Occurs when the instance is on the way to dispose
        /// </summary>
        event OnDisposeEventHandler OnDispose;

        /// <summary>
        /// Returns information the instance is already disposed
        /// </summary>
        bool IsDisposed { get; }

        /// <summary>
        /// Returns information the instance is currently in dispose operation
        /// </summary>
        bool IsCurrentlyDisposing { get; }

        /// <summary>
        /// Dispose the instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose open event subscriptions</param>
        void Dispose(bool disposeEventBinding);
    }
}
