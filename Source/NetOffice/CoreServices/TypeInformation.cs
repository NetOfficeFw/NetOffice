using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Represents a cache information about a com proxy and its contract and implementation in NetOffice
    /// </summary>
    [DebuggerDisplay("{Contract}")]
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
        /// <exception cref ="ArgumentException">contract/implementation or ids invalid</exception>
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
                throw new ArgumentException("Contract is not an interface.");
            if (implementation.IsInterface)
                throw new ArgumentException("Implementation must be a class.");
            if (!proxy.IsCOMObject)
                throw new ArgumentException("Proxy must be a com object.");
            if (componentId == Guid.Empty)
                throw new ArgumentException("Invalid component id.");
            if (typeId == Guid.Empty)
                throw new ArgumentException("Invalid type id.");

            Factory = factory;
            Contract = contract;
            Implementation = implementation;
            Proxy = proxy;
            ComponentId = componentId;
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
        /// Origin COM Component ID
        /// </summary>
        public Guid ComponentId { get; private set; }

        /// <summary>
        /// Origin COM Type ID
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

        /// <summary>
        /// Represents a System.String of the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("TypeInformation:{0},{1}", Contract.FullName, Implementation.FullName );
        }
    }
}
