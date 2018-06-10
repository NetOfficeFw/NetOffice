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
        /// <param name="factory">factory to create instances from</param>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <param name="proxy">proxy type</param>
        /// <param name="componentId">origin component id</param>
        /// <param name="typeId">origin type id</param>
        /// <exception cref ="ArgumentNullException">factory,contract, implementation or proxy is null</exception>
        /// <exception cref ="ArgumentException">contract is not an interface type</exception>
        internal TypeInformation(ITypeFactory factory, Type contract, Type implementation, Type proxy, Guid componentId, Guid typeId)
        {
            if (null == factory)
                throw new ArgumentNullException("factory");
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
        /// Factory to create instances from
        /// </summary>
        public ITypeFactory Factory { get; private set; }

        /// <summary>
        /// Contract Type
        /// </summary>
        public Type Contract { get; private set; }

        /// <summary>
        /// Implementation Type
        /// </summary>
        public Type Implementation { get; private set; }

        /// <summary>
        /// Proxy Type
        /// </summary>
        public Type Proxy { get; private set; }

        /// <summary>
        /// Origin COM Component Id
        /// </summary>
        public Guid ComponentId { get; private set; }

        /// <summary>
        /// Origin COM Type Id
        /// </summary>
        public Guid TypeId { get; private set; }

        /// <summary>
        /// Clones the instance
        /// </summary>
        /// <returns>newly created instance</returns>
        public TypeInformation Clone()
        {
            return new TypeInformation(Factory, Contract, Implementation, Proxy, ComponentId, TypeId);
        }
    }
}
