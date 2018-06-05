using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreSupport
{
    internal class TypeDictionary : List<TypeInformation>
    {
        public bool TryGetTypeInfo(string fullContractName, ref TypeInformation typeInfo)
        {
            foreach (var item in this)
            {
                if (fullContractName == item.Contract.FullName)
                {
                    typeInfo = item;
                    return true;
                }
            }
            return false;
        }

        public bool TryGetTypeInfo(Type contract, ref TypeInformation typeInfo)
        {
            foreach (var item in this)
            {
                if (contract == item.Contract)
                {
                    typeInfo = item;
                    return true;
                }
            }
            return false;
        }

        public bool TryGetProxyType(Type contract, ref Type proxy)
        {
            foreach (var item in this)
            {
                if (contract == item.Contract)
                {
                    proxy = item.Proxy;
                    return true;
                }
            }
            return false;
        }

        public void Add(Type contract, Type implementation, Type proxy)
        {
            Add(new TypeInformation(contract, implementation, proxy));
        }
    }
}
