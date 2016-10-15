using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;

namespace RegAddin.Common
{
    internal class AssemblyResolve
    {
        internal static Assembly Resolve(string name)
        {
            try
            {
                if (File.Exists(name))
                {
                    Assembly result = Assembly.LoadFile(name);
                    return result;
                }
                else if (!String.IsNullOrWhiteSpace(name))
                {
                    Assembly result = Assembly.Load(name);
                    return result;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
