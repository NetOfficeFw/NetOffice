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
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <param name="parentAssembly">parent assembly</param>
        internal DependentAssembly(string name, Assembly parentAssembly)
        {
            Name = name;
            ParentAssembly = parentAssembly;
        }

        /// <summary>
        /// Name of the assembly
        /// </summary>
        public string Name;

        /// <summary>
        /// Parent assembly
        /// </summary>
        public Assembly ParentAssembly;
    }
}
