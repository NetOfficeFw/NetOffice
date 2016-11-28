using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    internal class InstanceTypeNameResolver
    {
        internal string GetComponentName(ICOMObject instance)
        {
            return instance.InstanceType.Namespace;
        }

        internal string GetFriendlyInstanceName(ICOMObject instance)
        {
            string result = instance.InstanceType.FullName;
            result = result.Replace("NetOffice.", String.Empty);
            result = result.Replace("Api", String.Empty);
            return result;
        }
    }
}
