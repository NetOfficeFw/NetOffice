using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Represents various dispose possibilities to free child instances
    /// </summary>
    public interface ICOMObjectTableDisposable
    {
        /// <summary>
        /// Dispose all child instances
        /// </summary>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        void DisposeChildInstances();

        /// <summary>
        /// Dispose all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose open event subscriptions</param>
        /// <exception cref="COMDisposeException">An unexpected error occurs.</exception>
        void DisposeChildInstances(bool disposeEventBinding);
    }
}