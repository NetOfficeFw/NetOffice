using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Dependent assembly description
    /// </summary>
    internal struct DependentAssembly
    {
        internal DependentAssembly(string name, Assembly parentAssembly)
        {
            if (String.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (null == parentAssembly)
                throw new ArgumentNullException("parentAssembly");
            Name = name;
            ParentAssembly = parentAssembly;
        }

        public string Name;
        public Assembly ParentAssembly;
    }
}
