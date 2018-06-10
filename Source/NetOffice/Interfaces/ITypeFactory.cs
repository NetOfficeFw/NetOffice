using System;
using System.Collections.Generic;
using System.Reflection;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// NetOffice API Assembly Default Factory
    /// </summary>
    /// <remarks>Implementation is created directly by NetOffice Core and need to have a public ctor without arguments. Instance want create by name build by assembly default namespace,"Tools.Expose","TypeFactory"</remarks>
    public interface ITypeFactory
    {
        /// <summary>
        /// Simple name of the factory
        /// </summary>
        string FactoryName { get; }

        /// <summary>
        /// Default namespace of the factory
        /// </summary>
        string FactoryNamespace { get; }

        /// <summary>
        /// Guid of the COM component which represents the NetOfficeApi assembly
        /// </summary>
        Guid ComponentID { get; }

        /// <summary>
        /// Native API assembly
        /// </summary>
        Assembly Assembly { get; }

        /// <summary>
        /// NetOffice Assembly attribute
        /// </summary>
        /// <remarks>NetOffice Core want check the version for compatibility while initialize</remarks>
        NetOfficeAssemblyAttribute AssemblyAttribute { get; }

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        string[] Dependencies { get; }

        /// <summary>
        /// Exported/Public Types
        /// </summary>
        Type[] ExportedTypes { get; }

        /// <summary>
        /// Registered NetOffice Contract,Implementation Types
        /// </summary>
        /// <returns>contract,implementation pairs</returns>
        IEnumerable<KeyValuePair<Type, Type>> FactoryTypes();

        /// <summary>
        /// Returns info a class with given name exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="typeName">target type name</param>
        /// <returns>true if exists, otherwise false</returns>
        /// <exception cref= "ArgumentNullException">typeName is null(Nothing in Visual Basic) or empty</exception>
        bool ContainsType(string typeName);

        /// <summary>
        /// Returns info a class with given type exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="type">target type</param>
        /// <returns>true if exists, otherwise false</returns>
        /// <exception cref= "ArgumentNullException">type is null(Nothing in Visual Basic)</exception>
        bool ContainsType(Type type);

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
        /// <exception cref ="ArgumentNullException">contract is null</exception>
        bool Implementation(Type contract, ref Type implementation);

        /// <summary>
        /// Creates an instance of implementation type
        /// </summary>
        /// <param name="implementation"></param>
        /// <returns></returns>
        /// <exception cref ="ArgumentNullException">implementation is null</exception>
        /// <exception cref ="CreateFactoryInstanceException">unexepected error. see inner exception(s) for details</exception>
        ICOMObject CreateInstance(Type implementation);
    }
}
