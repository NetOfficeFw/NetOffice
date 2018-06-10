using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Loader;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Loaded Factories and known dependent assemblies
    /// </summary>
    public interface ICoreFactories
    {
        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        Core Parent { get; }

        /// <summary>
        /// Loaded Factories
        /// </summary>
        IEnumerable<ITypeFactory> Factories { get; }

        /// <summary>
        /// Known dependent NetOffice assemblies
        /// </summary>
        IEnumerable<DependentAssembly> Dependents { get; }
    }
}
