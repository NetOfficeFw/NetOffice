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
        /// Clone instance as target type of T
        /// </summary>
        /// <typeparam name="T">any other type to convert</typeparam>
        /// <returns>cloned instance</returns>
        /// <exception cref="CloneException">An unexpected error occurs.</exception>
        T To<T>() where T : class, ICOMObject;
    }
}