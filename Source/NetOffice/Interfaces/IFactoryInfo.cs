using System;
using System.Reflection;
using NetOffice.Attributes;

namespace NetOffice
{
    /// <summary>
    /// Informations about a NetOffice assembly
    /// </summary>
    public interface IFactoryInfo
    {
        /// <summary>
        /// Simple name of the assembly and type exporter
        /// </summary>
        string AssemblyName { get; }

        /// <summary>
        /// Namespace of the assembly
        /// </summary>
        string AssemblyNamespace { get; }

        /// <summary>
        /// Guid of component there represents the NetOfficeApi assembly
        /// </summary>
        Guid[] ComponentGuid { get; }

        /// <summary>
        /// Native API assembly
        /// </summary>
        Assembly Assembly { get; }

        /// <summary>
        /// Assembly attribute - Core want check the version for compatibility while initialize
        /// </summary>
        NetOfficeAssemblyAttribute AssemblyAttribute { get; }

        /// <summary>
        /// Returns info a class with given name exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="className">target class name</param>
        /// <returns>true if exists, otherwise false</returns>
        bool Contains(string className);

        /// <summary>
        /// Returns info a class with given type exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="type">target type</param>
        /// <returns>true if exists, otherwise false</returns>
        bool Contains(Type type);

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        string[] Dependencies { get; }

        /// <summary>
        /// Returns information the factory serves duck interfaces only
        /// </summary>
        bool IsDuck { get; }
    }
}
