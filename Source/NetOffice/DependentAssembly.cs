using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// dependent assembly description
    /// </summary>
    internal struct DependentAssembly
    {
        internal DependentAssembly(string name, Assembly parentAssembly)
        {
            Name = name;
            ParentAssembly = parentAssembly;
        }

        public string Name;
        public Assembly ParentAssembly;
    }
}
