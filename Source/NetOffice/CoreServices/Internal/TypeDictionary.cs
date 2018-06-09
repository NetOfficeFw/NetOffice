using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Represents the core type cache when resolving a wrapper type for a com proxy
    /// </summary>
    internal class TypeDictionary : List<TypeInformation>
    {
        #region Fields

        private object _thisLock = new object();

        #endregion

        #region Methods

        /// <summary>
        /// Adds new type info to the instance
        /// </summary>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <param name="proxy">proxy type</param>
        /// <param name="componentId">component id</param>
        /// <param name="typeId">type id</param>
        /// <exception cref ="ArgumentNullException">one or more arguments is null</exception>
        /// <exception cref ="ArgumentException">one or more arguments does not match</exception>
        /// <returns>newly created typeInformation</returns>
        internal TypeInformation Add(Type contract, Type implementation, Type proxy, Guid componentId, Guid typeId)
        {
            if (null == contract)
                throw new ArgumentNullException("contract");
            if (null == implementation)
                throw new ArgumentNullException("implementation");
            if (null == proxy)
                throw new ArgumentNullException("proxy");
#if DEBUG
            if (this.Any(e => e.Contract == contract))
                throw new ArgumentException("Duplicated contract in type cache detected.");
#endif

            TypeInformation result = null;

            lock (_thisLock)
            {
                if (!contract.IsInterface)
                    throw new ArgumentException("contract is not an interface.");
                if (implementation.IsInterface)
                    throw new ArgumentException("implementation must be a class.");
                if (!proxy.IsCOMObject)
                    throw new ArgumentException("proxy must be a com object.");

                result = new TypeInformation(contract, implementation, proxy, componentId, typeId);
                Add(result);
            }
            return result;
        }

        internal TypeInformation TryGetTypeInfo(Guid componentId, Guid typeId)
        {
            lock (_thisLock)
            {
                foreach (var item in this)
                {
                    if (componentId == item.ComponentId && typeId == item.TypeId)
                    {
                        return item;
                    }
                }
                return null;
            }
        }

        /// <summary>
        ///  Get type info by contract type
        /// </summary>
        /// <param name="contract">contract type</param>
        /// <returns>type info or null if not found</returns>
        /// <exception cref ="ArgumentNullException">contract is null</exception>
        /// <exception cref ="ArgumentException">contract is not an interface type</exception>
        internal TypeInformation TryGetTypeInfo(Type contract)
        {
            if (null == contract)
                throw new ArgumentNullException("contract");
            if (!contract.IsInterface)
                throw new ArgumentException("contract is not an interface.");

            lock (_thisLock)
            {
                foreach (var item in this)
                {
                    if (contract == item.Contract)
                    {
                        return item;
                    }
                }
                return null;
            }
        }

        ///// <summary>
        ///// Get type info by full qualified contract name
        ///// </summary>
        ///// <param name="fullContractName">full qualified contract name</param>
        ///// <param name="typeInfo">result or null(Nothing in Visual Basic)</param>
        ///// <returns>true if type info is delivered, otherwise false</returns>
        ///// <exception cref ="ArgumentNullException">fullContractName is null</exception>
        //internal bool TryGetTypeInfo(string fullContractName, ref TypeInformation typeInfo)
        //{
        //    if (String.IsNullOrWhiteSpace(fullContractName))
        //        throw new ArgumentNullException("fullContractName");

        //    lock (_thisLock)
        //    {
        //        foreach (var item in this)
        //        {
        //            if (fullContractName == item.Contract.FullName)
        //            {
        //                typeInfo = item;
        //                return true;
        //            }
        //        }
        //        return false;
        //    }
        //}

        ///// <summary>
        /////  Get type info by contract type
        ///// </summary>
        ///// <param name="contract">contract type</param>
        ///// <param name="typeInfo">result or null(Nothing in Visual Basic)</param>
        ///// <returns>true if type info is delivered, otherwise false</returns>
        ///// <exception cref ="ArgumentNullException">contract is null</exception>
        ///// <exception cref ="ArgumentException">contract is not an interface type</exception>
        //internal bool TryGetTypeInfo(Type contract, ref TypeInformation typeInfo)
        //{
        //    if (null == contract)
        //        throw new ArgumentNullException("contract");
        //    if (!contract.IsInterface)
        //        throw new ArgumentException("contract is not an interface.");

        //    lock (_thisLock)
        //    {
        //        foreach (var item in this)
        //        {
        //            if (contract == item.Contract)
        //            {
        //                typeInfo = item;
        //                return true;
        //            }
        //        }
        //        return false;
        //    }          
        //}

        ///// <summary>
        /////  Get proxy type by contract type
        ///// </summary>
        ///// <param name="contract">contract type</param>
        ///// <param name="proxy">result or null(Nothing in Visual Basic)</param>
        ///// <returns>true if type info is delivered, otherwise false</returns>
        ///// <exception cref ="ArgumentNullException">contract is null</exception>
        ///// <exception cref ="ArgumentException">contract is not an interface type</exception>
        //internal bool TryGetProxyType(Type contract, ref Type proxy)
        //{
        //    if (null == contract)
        //        throw new ArgumentNullException("contract");
        //    if (!contract.IsInterface)
        //        throw new ArgumentException("contract is not an interface.");

        //    lock (_thisLock)
        //    {
        //        foreach (var item in this)
        //        {
        //            if (contract == item.Contract)
        //            {
        //                proxy = item.Proxy;
        //                return true;
        //            }
        //        }
        //        return false;
        //    }
        //}



        /// <summary>
        /// Creates an enumerable copy
        /// </summary>
        /// <returns>newly created copy</returns>
        internal IEnumerable<TypeInformation> ToEnumerable()
        {
            lock (_thisLock)
            {
                TypeInformation[] result = new TypeInformation[Count];
                for (int i = 0; i < Count; i++)
                {
                    result[i] = this[i].Clone();
                }
                return result;
            }
        }

        #endregion
    }
}
