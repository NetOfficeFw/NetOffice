using System;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal static class AttributeReflection
    {       
        internal static IEnumerable<object> GetCustomClassAttributes(Type item)
        {
            List<object> result = new List<object>();
            Type type = item;
            while (null != type)
            {
                object[] attributes = type.GetCustomAttributes(false);
                foreach (var attribute in attributes)
                    result.Add(attribute);
                if (null != type.BaseType && type.BaseType.FullName == "System.Object")
                    break;
                type = type.BaseType;
            }

            return result;
        }

        internal static bool ComVisibleAttributeExists(IEnumerable<object> customAttributes)
        {
            foreach (object item in customAttributes)
            {
                ComVisibleAttribute attribute = item as ComVisibleAttribute;
                if (null != attribute && attribute.Value)
                    return true;
            }
            return false;
        }
        
        internal static bool GuidAttributeExists(IEnumerable<object> customAttributes)
        {
            foreach (object item in customAttributes)
            {
                GuidAttribute attribute = item as GuidAttribute;
                if (null != attribute)
                    return true;
            }
            return false;
        }
  
        internal static bool ProgIdAttributeExists(IEnumerable<object> customAttributes)
        {
            foreach (object item in customAttributes)
            {
                ProgIdAttribute attribute = item as ProgIdAttribute;
                if (null != attribute)
                    return true;
            }
            return false;
        }

        internal static bool ClassInterfaceAttributeExists(IEnumerable<object> customAttributes)
        {
            foreach (object item in customAttributes)
            {
                ClassInterfaceAttribute attribute = item as ClassInterfaceAttribute;
                if (null != attribute)
                    return true;
            }
            return false;
        }

        internal static bool AttributeExists<T>(IEnumerable<object> customAttributes) where T : System.Attribute
        {
            foreach (object item in customAttributes)
            {
                T attribute = item as T;
                if (null != attribute)
                    return true;
            }
            return false;
        }

        internal static T GetAttribute<T>(IEnumerable<object> customAttributes) where T : System.Attribute
        {
            foreach (object item in customAttributes)
            {
                T attribute = item as T;
                if (null != attribute)
                    return attribute;
            }
            throw new ArgumentException("Unable to find attribute");
        }
    }
}
