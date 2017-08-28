using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// OnDispose Event Arguments
    /// </summary>
    public class OnDisposeEventArgs : EventArgs
    {
        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        /// <param name="sender">the target COM object</param>
        internal OnDisposeEventArgs(ICOMObject sender)
        {
            Sender = sender;
        }

        /// <summary>
        /// Target COM object
        /// </summary>
        public ICOMObject Sender { get; private set; }

        /// <summary>
        /// Skip flag, you can cancel the operation if you want
        /// </summary>
        public bool Cancel { get; set; }
    }

    /// <summary>
    /// EventHandler delegate for ICOMObjectDisposable.OnDispose
    /// </summary>
    /// <param name="eventArgs">dispose arguments</param>
    public delegate void OnDisposeEventHandler(OnDisposeEventArgs eventArgs);

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
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        void Dispose(bool disposeEventBinding);
    }
}