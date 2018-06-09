using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Represents a cache information about a com proxy and its contract and implementation in NetOffice
    /// </summary>
    public class TypeInformation
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <param name="proxy">proxy type</param>
        /// <param name="componentId">origin component id</param>
        /// <param name="typeId">origin type id</param>
        /// <exception cref ="ArgumentNullException">contract is null</exception>
        /// <exception cref ="ArgumentException">contract is not an interface type</exception>
        internal TypeInformation(Type contract, Type implementation, Type proxy, Guid componentId, Guid typeId)
        {
            if (null == contract)
                throw new ArgumentNullException("contract");
            if (null == implementation)
                throw new ArgumentNullException("implementation");
            if (null == proxy)
                throw new ArgumentNullException("proxy");

            if (!contract.IsInterface)
                throw new ArgumentException("contract is not an interface.");
            if (implementation.IsInterface)
                throw new ArgumentException("implementation must be a class.");
            if (!proxy.IsCOMObject)
                throw new ArgumentException("proxy must be com a object.");

            Contract = contract;
            Implementation = implementation;
            Proxy = proxy;
            TypeId = typeId;
        }

        /// <summary>
        /// Clones the instance
        /// </summary>
        /// <returns>newly created instance</returns>
        internal TypeInformation Clone()
        {
            return new TypeInformation(Contract, Implementation, Proxy, ComponentId, TypeId);
        }

        /// <summary>
        /// Contract Type
        /// </summary>
        internal Type Contract { get; private set; }

        /// <summary>
        /// Implementation Type
        /// </summary>
        internal Type Implementation { get; private set; }

        /// <summary>
        /// Proxy Type
        /// </summary>
        internal Type Proxy { get; private set; }

        /// <summary>
        /// Origin COM Component Id
        /// </summary>
        internal Guid ComponentId { get; private set; }

        /// <summary>
        /// Origin COM Type Id
        /// </summary>
        internal Guid TypeId { get; private set; }
    }
}
