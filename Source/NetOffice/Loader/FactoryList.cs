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
    public class FactoryList : List<IFactoryInfo>
    {
        /// <summary>
        /// Check for loaded assembly in factory list
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <returns>true if exists, otherwise false</returns>
        public bool Contains(string name)
        {
            if (String.IsNullOrWhiteSpace(name))
                return false;

            if (name.EndsWith(".dll", StringComparison.InvariantCultureIgnoreCase))
                name = name.Substring(0, name.Length - 4);

            foreach (IFactoryInfo item in this)
            {
                if (item.AssemblyName.StartsWith(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Returns implementation from contract
        /// </summary>
        /// <param name="contractType">target contract</param>
        /// <returns>implementation type</returns>
        /// <exception cref ="ArgumentNullException">argument is null</exception>
        /// <exception cref ="FactoryException">unexpected type load error</exception>
        public Type GetImplementationType(Type contractType)
        {
            if (null == contractType)
                throw new ArgumentNullException("contractType");

            try
            {
                string contractTypeNamespace = contractType.Namespace;
                var item = this.FirstOrDefault(e => e.AssemblyNamespace == contractTypeNamespace);
                string target = contractTypeNamespace + ".Behind." + contractType.Name;
                Type implementationResult = item.Assembly.GetType(target, true);
                var attribute = implementationResult.GetCustomAttribute<HasInteropCompatibilityClass>();
                if (null != attribute)
                    implementationResult = attribute.Value;
                return implementationResult;
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

        /// <summary>
        /// Returns contract and implementation type by contract name
        /// </summary>
        /// <param name="contractTypeNamespace">contract name space</param>
        /// <param name="contractTypeName">contract non-fullqualified name</param>
        /// <param name="contract">contract result</param>
        /// <param name="implementation">implementation result</param>
        /// <exception cref ="ArgumentNullException">argument is null or empty whitespace</exception>
        /// <exception cref ="FactoryException">unexpected type load error</exception>
        public void GetContractAndImplementationType(string contractTypeNamespace, string contractTypeName, ref Type contract, ref Type implementation)
        {
            if (String.IsNullOrWhiteSpace(contractTypeNamespace))
                throw new ArgumentNullException("contractTypeNamespace");
            if (String.IsNullOrWhiteSpace(contractTypeName))
                throw new ArgumentNullException("contractTypeName");

            try
            {
                var item = this.FirstOrDefault(e => e.AssemblyNamespace == contractTypeNamespace);
                string contractTarget = contractTypeNamespace + "." + contractTypeName;
                string implementationTarget = contractTypeNamespace + ".Behind." + contractTypeName;

                Type contractResult = item.Assembly.GetType(contractTarget, true);
                Type implementationResult = item.Assembly.GetType(contractTarget, true);

                var attribute = implementationResult.GetCustomAttribute<HasInteropCompatibilityClass>();
                if (null != attribute)
                    implementationResult = attribute.Value;

                contract = contractResult;
                implementation = implementationResult;
            }
            catch (TypeLoadException exception)
            {
                throw new FactoryException(String.Format("Unable to load contract or implementation type: {0}.{1}.", contractTypeNamespace, contractTypeName), exception);
            }
            catch (Exception exception)
            {
                throw new FactoryException(String.Format("Unexcepted type load error(2): {0}.{1}.", contractTypeNamespace, contractTypeName), exception);
            }
        }
    }
}
