using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreSupport
{
    internal class TypeInformation
    {
        internal TypeInformation(Type contract, Type implementation, Type proxy)
        {
            if (null == contract)
                throw new ArgumentNullException("contract");
            if (null == implementation)
                throw new ArgumentNullException("implementation");
            if (null == proxy)
                throw new ArgumentNullException("proxy");

            Contract = contract;
            Implementation = implementation;
            Proxy = proxy;
        }

        public Type Contract { get; private set; }

        public Type Implementation { get; private set; }

        public Type Proxy { get; private set; }
    }
}
