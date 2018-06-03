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
            string rootSpace = contractType.Namespace;
            var item = this.FirstOrDefault(e => e.AssemblyNamespace == rootSpace);
            string target = rootSpace + ".Behind." + contractType.Name;
            Type result = item.Assembly.GetType(target, true);
            var attribute = result.GetCustomAttribute<HasInteropCompatibilityClass>();
            if (null != attribute)
                result = attribute.Value;
            return result;
        }

        public bool GetContractAndImplementationType(string name, ref Type contract, ref Type implementation, bool throwException = false)
        {
            return false;
        }
    }
}
