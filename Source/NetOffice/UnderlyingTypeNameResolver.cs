using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    internal class UnderlyingTypeNameResolver
    {
        internal string GetComponentName(ICOMObject instance)
        {
            return null != instance ? TypeDescriptor.GetComponentName(instance.UnderlyingObject) : String.Empty;
        }

        internal string GetClassName(ICOMObject instance)
        {
            return null != instance ? TypeDescriptor.GetClassName(instance.UnderlyingObject) : String.Empty;
        }

        internal string GetFriendlyClassName(ICOMObject instance, string className)
        {
            string fullname = null != className ? className : GetFriendlyClassName(instance);
            return fullname;
        }

        internal string GetFriendlyClassName(ICOMObject instance)
        {
            string fullName = instance.UnderlyingType.FullName;
            fullName = fullName.Replace("Microsoft", String.Empty).Replace("Interop", String.Empty).Replace("..", ".");
            return fullName;
        }
    }
}
