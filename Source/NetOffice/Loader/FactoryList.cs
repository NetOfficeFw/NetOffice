using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.Loader
{
    /// <summary>
    /// Contains loaded factory informations
    /// </summary>
    public class FactoryList 
    {
        private static string _dllExtension = ".dll";
        private List<ITypeFactory> _items = new List<ITypeFactory>();

        /// <summary>
        /// Check for loaded assembly in factory list
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <returns>true if exists, otherwise false</returns>
        public bool Contains(string name)
        {
            if (String.IsNullOrWhiteSpace(name))
                return false;

            if (name.EndsWith(_dllExtension, StringComparison.InvariantCultureIgnoreCase))
                name = name.Substring(0, name.Length - _dllExtension.Length);

            foreach (ITypeFactory item in _items)
            {
                if (item.FactoryName.StartsWith(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Returns type factory from contact
        /// </summary>
        /// <param name="contractType">target contract type</param>
        /// <returns>corresponding type factory</returns>
        /// <exception cref="ArgumentNullException">contractType is null</exception>
        /// <exception cref="InvalidOperationException">unable to find type factory</exception>
        public ITypeFactory GetTypeFactory(Type contractType)
        {
            if (null == contractType)
                throw new ArgumentNullException("contractType");
            string contractTypeNamespace = contractType.Namespace;
            ITypeFactory item = _items.First(e => e.FactoryNamespace == contractTypeNamespace);
            return item;
        }

        /// <summary>
        /// Returns contract and implementation type from given factory
        /// </summary>
        /// <param name="factory">factory to look into</param>
        /// <param name="typeId">target type id</param>
        /// <param name="contract">corresponding contract</param>
        /// <param name="implementation">corresponding implementation</param>
        /// <returns>true if both filled, otherwise false</returns>
        public bool GetContractAndImplementationType(ITypeFactory factory, Guid typeId, ref Type contract, ref Type implementation)
        {
            bool result = false;
            result = factory.ContractAndImplementation(typeId, ref contract, ref implementation);
            if (result)
            {
                var coClass = contract.GetCustomAttribute<CoClassSourceAttribute>();
                if (null != coClass)
                {
                    contract = coClass.Value;
                    if (!factory.Implementation(contract, ref implementation))
                        return false;
                }

                var attribute = implementation.GetCustomAttribute<HasInteropCompatibilityClassAttribute>();
                if (null != attribute)
                    implementation = attribute.Value;
            }
            return result;
        }

        /// <summary>
        /// Returns implementation from contract
        /// </summary>
        /// <param name="contractType">target contract</param>
        /// <param name="throwException">throw exception if failed to resolve</param>
        /// <returns>implementation type</returns>
        /// <exception cref ="ArgumentNullException">argument is null</exception>
        /// <exception cref ="FactoryException">unexpected type load error</exception>
        /// <exception cref ="ArgumentException">unable to find result and exception should thrown</exception>
        public Type GetImplementationType(Type contractType, bool throwException = true)
        {
            ITypeFactory factory = null;
            return GetImplementationType(contractType, ref factory, throwException);
        }

        /// <summary>
        /// Returns implementation from contract
        /// </summary>
        /// <param name="contractType">target contract</param>
        /// <param name="factory">corresponding factory</param>
        /// <param name="throwException">throw exception if failed to resolve</param>
        /// <returns>implementation type</returns>
        /// <exception cref ="ArgumentNullException">argument is null</exception>
        /// <exception cref ="FactoryException">unexpected type load error</exception>
        /// <exception cref ="ArgumentException">unable to find result and exception should thrown</exception>
        public Type GetImplementationType(Type contractType, ref ITypeFactory factory, bool throwException = true)
        {
            if (null == contractType)
                throw new ArgumentNullException("contractType");

            Type result = null;
            try
            {
                string contractTypeNamespace = contractType.Namespace;
                factory = _items.FirstOrDefault(e => e.FactoryNamespace == contractTypeNamespace);
                if (null != factory)
                    factory.Implementation(contractType, ref result);

                if (null != result)
                {
                    var attribute = result.GetCustomAttribute<HasInteropCompatibilityClassAttribute>();
                    if (null != attribute)
                        result = attribute.Value;
                }
                else if (throwException)
                {
                    throw new ArgumentException("Unable to find implementation.");
                }

                return result;
            }
            catch (ArgumentNullException)
            {
                throw;
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (TypeLoadException exception)
            {
                throw new FactoryException(String.Format("Unable to load implementation type: {0}.", contractType.FullName), exception);
            }
            catch (Exception exception)
            {
                throw new FactoryException(String.Format("Unexcepted type load error(1): {0}.", contractType.FullName), exception);
            }
        }
    }
}
