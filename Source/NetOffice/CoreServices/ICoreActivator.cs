using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides ICOMObject Activation Services
    /// </summary>
    public interface ICoreActivator
    {
        /// <summary>
        /// Occours when a new COMObject instance has been created
        /// </summary>
        event OnCreateInstanceEventHandler CreateInstance;

        /// <summary>
        /// Occurs when a new COMDynamicObject instance should be created
        /// </summary>
        event OnCreateCOMDynamicEventHandler CreateDynamicInstance;

        /// <summary>
        /// Occurs when a new COMProxyShare instance should be created
        /// </summary>
        event OnCreateProxyShareEventHandler CreateProxyShare;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }

        /// <summary>
        /// Registered custom types
        /// </summary>
        IEnumerable<KeyValuePair<Type, Type>> RegisteredTypes { get; }

        /// <summary>
        /// Add a custom type
        /// </summary>
        /// <param name="contract">target contract</param>
        /// <param name="implementation">custom implementation</param>
        void RegisterType(Type contract, Type implementation);

        /// <summary>
        /// Remove a custom type
        /// </summary>
        /// <param name="contract">target contract</param>
        /// <returns>true, if removed, otherwise false</returns>
        bool UnRegisterType(Type contract);        
    }
}
