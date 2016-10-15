using System;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Info about a NetOffice assembly
    /// </summary>
    public interface IFactoryInfo
    {
        /// <summary>
        /// Namespace of assembly
        /// </summary>
        string AssemblyNamespace { get; }

        /// <summary>
        /// Guid of component there represents the NetOfficeApi assembly
        /// </summary>
        Guid[] ComponentGuid { get; }

        /// <summary>
        /// Assembly info of NetOfficeApi assembly
        /// </summary>
        Assembly Assembly { get; }

        /// <summary>
        /// Returns info a class with given name exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        bool Contains(string className);

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        string[] Dependencies { get; }
    }
}
