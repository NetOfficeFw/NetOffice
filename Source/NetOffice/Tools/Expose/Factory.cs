using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.Tools.Expose
{
    /// <summary>
    /// Api Assembly Factory Base Type
    /// </summary>
    public abstract class Factory : ITypeFactory
    {
        #region Fields

        private string _factoryName;
        private Assembly _assembly;
        private NetOfficeAssemblyAttribute _assemblyAttribute;
        private Type[] _exportedTypes;
        private Dictionary<Type, Type> _factoryTypes;

        #endregion

        #region ITypeFactory

        /// <summary>
        /// Simple name of the factory
        /// </summary>
        public virtual string FactoryName
        {
            get
            {
                if (null == _factoryName)
                {
                    _factoryName = Assembly.GetName().Name;
                }
                return _factoryName;
            }
        }

        /// <summary>
        /// Default namespace of the factory
        /// </summary>
        public abstract string FactoryNamespace { get; }

        /// <summary>
        /// Guid of the COM component which represents the NetOfficeApi assembly
        /// </summary>
        public abstract Guid ComponentID { get; }

        /// <summary>
        /// Native API assembly
        /// </summary>
        public virtual Assembly Assembly
        {
            get
            {
                if (null == _assembly)
                {
                    _assembly = GetType().Assembly;
                }
                return _assembly;
            }
        }

        /// <summary>
        /// NetOffice Assembly attribute
        /// </summary>
        /// <remarks>NetOffice Core want check the version for compatibility while initialize</remarks>
        public NetOfficeAssemblyAttribute AssemblyAttribute
        {
            get
            {
                if (null == _assemblyAttribute)
                {
                    _assemblyAttribute = Assembly.GetCustomAttributes(typeof(NetOfficeAssemblyAttribute), true)[0] as NetOfficeAssemblyAttribute;
                }
                return _assemblyAttribute;
            }
        }

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        public abstract string[] Dependencies { get; }

        /// <summary>
        /// Exported/Public Types
        /// </summary>
        public virtual Type[] ExportedTypes
        {
            get
            {
                if (null == _exportedTypes)
                    _exportedTypes = Assembly.GetExportedTypes();
                return _exportedTypes;
            }
        }

        /// <summary>
        /// Registered NetOffice Contract,Implementation Types
        /// </summary>
        /// <returns>contract,implementation pairs</returns>
        public virtual IEnumerable<KeyValuePair<Type, Type>> FactoryTypes()
        {
            CreateFactoryTypes();
            return _factoryTypes.ToArray();
        }

        /// <summary>
        /// Returns info a class with given name exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="typeName">target type name</param>
        /// <returns>true if exists, otherwise false</returns>
        /// <exception cref= "ArgumentNullException">typeName is null(Nothing in Visual Basic) or empty</exception>
        public virtual bool ContainsType(string typeName)
        {
            if (String.IsNullOrWhiteSpace(typeName))
                throw new ArgumentNullException("typeName");

            return ExportedTypes.Any(e => e.Name.EndsWith(typeName, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Returns info a class with given type exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="type">target type</param>
        /// <returns>true if exists, otherwise false</returns>
        /// <exception cref= "ArgumentNullException">type is null(Nothing in Visual Basic)</exception>
        public virtual bool ContainsType(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");

            return ExportedTypes.Any(e => e == type);
        }

        /// <summary>
        /// Returns contract and implementation type by COM type id
        /// </summary>
        /// <param name="typeId">target type id</param>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <returns>true if both filled, otherwise false</returns>
        /// <exception cref="TypeLoadException">failed to recieve implementation for a contract</exception>
        public virtual bool ContractAndImplementation(Guid typeId, ref Type contract, ref Type implementation)
        {
            CreateFactoryTypes();
            foreach (var item in _factoryTypes)
            {
                if (item.Key.GetCustomAttribute<TypeIdAttribute>().Value == typeId)
                {
                    contract = item.Key;
                    implementation = item.Value;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Returns an implementation by its contract
        /// </summary>
        /// <param name="contract">contract type</param>
        /// <param name="implementation">implementation type</param>
        /// <returns>true if filled, otherwise false</returns>
        /// <exception cref ="ArgumentNullException">contract is null</exception>
        /// <exception cref="TypeLoadException">failed to recieve implementation for a contract</exception>
        public virtual bool Implementation(Type contract, ref Type implementation)
        {
            if (null == contract)
                throw new ArgumentNullException("contract");
            CreateFactoryTypes();
            bool result =  _factoryTypes.TryGetValue(contract, out implementation);
            return result;
        }

        /// <summary>
        /// Creates an instance of implementation type
        /// </summary>
        /// <param name="implementation"></param>
        /// <returns></returns>
        /// <exception cref ="ArgumentNullException">implementation is null</exception>
        /// <exception cref ="CreateFactoryInstanceException">unexepected error. see inner exception(s) for details</exception>
        /// <exception cref="TypeLoadException">failed to recieve implementation for a contract</exception>
        public virtual ICOMObject CreateInstance(Type implementation)
        {
            if (null == implementation)
                throw new ArgumentNullException("implementation");
            try
            {
                if (!ContainsType(implementation))
                    throw new ArgumentException("Type is not managed by factory");

                if(implementation.HasCustomAttribute<InteropCompatibilityClassAttribute>())
                    return (ICOMObject)Activator.CreateInstance(implementation, NetOffice.Callers.InteropCompatibilityClassCreateMode.FromActivator);
                else
                    return (ICOMObject)Activator.CreateInstance(implementation);
            }
            catch (Exception exception)
            {
                throw new CreateFactoryInstanceException(exception);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates Factory Types Dictionary
        /// </summary>
        /// <exception cref="TypeLoadException">failed to recieve implementation for a contract</exception>
        protected internal virtual void CreateFactoryTypes()
        {
            if (null == _factoryTypes)
            {
                _factoryTypes = new Dictionary<Type, Type>();
                var contracts = ExportedTypes.Where(e => e.IsInterface
                                && e.Namespace == FactoryNamespace
                                && false == e.HasCustomAttribute<SyntaxBypassAttribute>());
                foreach (var contract in contracts)
                {
                    var attribute = contract.GetCustomAttribute<EntityTypeAttribute>();
                    if (null != attribute && 
                        (attribute.Type == EntityType.IsDispatchInterface ||
                        attribute.Type == EntityType.IsCoClass ||
                        attribute.Type == EntityType.IsInterface ||
                        attribute.Type == EntityType.IsNativeInterfaceCaller
                        ))
                    {
                        var implementation = Assembly.GetType(contract.Namespace + ".Behind." + contract.Name, true);
                        _factoryTypes.Add(contract, implementation);
                    }
                }
            }
        }

        #endregion
    }
}
