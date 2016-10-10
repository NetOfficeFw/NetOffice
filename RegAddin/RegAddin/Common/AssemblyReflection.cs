using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RegAddin.Common
{
    internal class AssemblyReflection
    {
        internal static IEnumerable<object> GetCustomAssemblyAttributes(Assembly assembly)
        {
            return assembly.GetCustomAttributes(true);
        }

        internal static bool AssemblyIsComVisible(Assembly assembly, IEnumerable<object> attributes)
        {
            if (null == attributes)
                attributes = assembly.GetCustomAttributes(true);

            foreach (object item in attributes)
            {
                ComVisibleAttribute comVisible = item as ComVisibleAttribute;
                if (null != comVisible && comVisible.Value == true)
                    return true;
            }
            return false;
        }
    }
}
