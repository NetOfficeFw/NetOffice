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
        /// Guid of the COM component which represents the NetOfficeApi assembly
        /// </summary>
        Guid ComponentGuid { get; }

        /// <summary>
        /// Native API assembly
        /// </summary>
        Assembly Assembly { get; }

        /// <summary>
        /// Assembly attribute - Core want check the version for compatibility while initialize
        /// </summary>
        NetOfficeAssemblyAttribute AssemblyAttribute { get; }

        /// <summary>
        /// Exported/Public Types
        /// </summary>
        Type[] ExportedTypes { get; }

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        string[] Dependencies { get; }

        /// <summary>
        /// Returns information the factory serves duck interfaces only
        /// </summary>
        bool IsDuck { get; }

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
        /// Returns contract and implementation type by COM type id
        /// </summary>
        /// <param name="typeId">target type id</param>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <returns>true if both filled, otherwise false</returns>
        bool ContractAndImplementation(Guid typeId, ref Type contract, ref Type implementation);

        /// <summary>
        /// Returns an implementation by its contract
        /// </summary>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <returns>true if filled, otherwise false</returns>
        bool Implementation(Type contract, ref Type implementation);
    }
}
