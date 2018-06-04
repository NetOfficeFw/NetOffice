using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.Loader
{
    /// <summary>
    /// Contains loaded factory informations
    /// </summary>
    public class FactoryList: List<IFactoryInfo>
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

        public Type GetImplementationType(Type contractType)
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

        public void GetContractAndImplementationType(string contractTypeNamespace, string contractTypeName, ref Type contract, ref Type implementation)
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
    }
}
