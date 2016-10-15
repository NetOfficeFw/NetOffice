using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal static class AddinClassReflection
    {
        internal static bool IsValidAddinClass(IEnumerable<object> customAttributes, TypeAttributes typeAttributes)
        {
            TypeAttributes isPublic = typeAttributes & TypeAttributes.Public;
            if (isPublic != TypeAttributes.Public)
                return false;
            
            if (false == AttributeReflection.ComVisibleAttributeExists(customAttributes) ||
                false == AttributeReflection.GuidAttributeExists(customAttributes) ||
                false == AttributeReflection.ProgIdAttributeExists(customAttributes))
                return false;
            else
                return true;
        }

        internal static bool IsValidAddinClass(IEnumerable<object> customAttributes)
        {
            if (false == AttributeReflection.ComVisibleAttributeExists(customAttributes) ||
                false == AttributeReflection.GuidAttributeExists(customAttributes) ||
                false == AttributeReflection.ProgIdAttributeExists(customAttributes))
                return false;
            else
                return true;
        }
    }
}
