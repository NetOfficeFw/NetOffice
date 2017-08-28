using System;

namespace NetOffice
{
    /// <summary>
    /// Represents a managed/wrapped COM proxy implementation 
    /// </summary>
    public interface ICOMObject : ICOMObjectProxy, ICOMObjectDisposable, ICOMObjectTable, ICOMObjectTableDisposable, ICOMObjectEvents, ICOMObjectAvaility, ICloneable
    {
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
    }
}