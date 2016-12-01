using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Represents a managed/wrapped COM proxy implementation 
    /// </summary>
    public interface ICOMObject : ICOMObjectTable, ICOMObjectProxy, ICOMObjectTableDisposable, ICOMObjectEvents, ICOMObjectDisposable
    {
        /// <summary>
        /// Returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <returns>true if available, otherwise false</returns>
        bool EntityIsAvailable(string name);

        /// <summary>
        /// Returns information the proxy provides a method or property with given name at runtime
        /// </summary>
        /// <param name="name">name of the enitity</param>
        /// <param name="searchType">indicate the kind of pr0operty</param>
        /// <returns>true if available, otherwise false</returns>
        bool EntityIsAvailable(string name, SupportEntityType searchType);

        /// <summary>
        /// The associated console
        /// </summary>
        DebugConsole Console { get; }

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
    }
}
