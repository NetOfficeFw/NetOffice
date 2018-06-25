using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Represents a managed/wrapped COM proxy implementation
    /// </summary>
    public interface ICOMObject : ICOMObjectProxy, ICOMObjectDisposable, ICOMObjectTable, ICOMObjectTableDisposable, ICOMObjectEvents, ICOMObjectAvaility, ICloneable
    {
        /// <summary>
        /// Monitor Lock
        /// </summary>
        object SyncRoot { get; }

        /// <summary>
        /// The associated factory
        /// </summary>
        Core Factory { get; }

        /// <summary>
        /// The associated invoker
        /// </summary>
        Invoker Invoker { get; }

        /// <summary>
        /// The associated settings
        /// </summary>
        Settings Settings { get; }

        /// <summary>
        /// The associated console
        /// </summary>
        DebugConsole Console { get; }

        /// <summary>
        /// Clone instance as T
        /// </summary>
        /// <typeparam name="T">type to convert</typeparam>
        /// <returns>cloned instance</returns>
        /// <exception cref="CloneException">An unexpected error occurs.</exception>
        T To<T>() where T : class, ICOMObject;

        /// <summary>
        /// Determines whether two ICOMObject instances pointing to the same remote server instance
        /// </summary>
        /// <param name="obj">target instance to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        bool EqualsOnServer(object obj);
    }
}