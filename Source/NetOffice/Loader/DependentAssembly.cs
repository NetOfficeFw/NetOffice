using System;
using System.Reflection;

namespace NetOffice.Loader
{
    /// <summary>
    /// Dependent assembly description
    /// </summary>
    internal struct DependentAssembly
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">Name of the dependent assembly</param>
        /// <param name="parentAssembly">assembly that is required from</param>
        internal DependentAssembly(string name, Assembly parentAssembly)
        {
            if (String.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("name");
            if (null == parentAssembly)
                throw new ArgumentNullException("parentAssembly");
            Name = name;
            ParentAssembly = parentAssembly;
        }

        /// <summary>
        /// Name of the dependent assembly
        /// </summary>
        public string Name;

        /// <summary>
        /// Assembly that is required from
        /// </summary>
        public Assembly ParentAssembly;
    }
}