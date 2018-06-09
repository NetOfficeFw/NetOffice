using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using NetOffice;
using NetOffice.Attributes;

namespace VBIDEApi.Utils
{
    #pragma warning disable
    /// <summary>
    /// Necessary factory info, used from NetOffice.Core while Initialize()
    /// </summary>
    public class ProjectInfo : IFactoryInfo
    {
        #region Fields

        private string    _name;
        private string    _namespace     = "NetOffice.VBIDEApi";
        private Guid      _componentGuid = new Guid("0002E157-0000-0000-C000-000000000046");
        private Assembly  _assembly;
        private NetOfficeAssemblyAttribute _assemblyAttribute;
        private Type[]	  _exportedTypes;
		private string[]  _dependents;
        private Dictionary<Type, Type> _types;

        #endregion

        #region Ctor

        public ProjectInfo()
        {
            _assembly = typeof(ProjectInfo).Assembly;
            _assemblyAttribute = _assembly.GetCustomAttributes(typeof(NetOfficeAssemblyAttribute), true)[0] as NetOfficeAssemblyAttribute;
            _name = _assembly.GetName().Name;
        }

        #endregion

        #region IFactoryInfo

        public string AssemblyName
        {
            get
            {
                return _name;
            }
        }

        public string AssemblyNamespace
        {
            get
            {
                return _namespace;
            }
        }

        public Guid ComponentGuid
        {
            get
            {
                return _componentGuid;
            }
        }

        public Assembly Assembly
        {
            get
            {
                return _assembly;
            }
        }

        public NetOfficeAssemblyAttribute AssemblyAttribute
        {
            get
            {
                return _assemblyAttribute;
            }
        }

        public string[] Dependencies
        {
            get
            {
				if(null == _dependents)
					_dependents = new string[]{"OfficeApi.dll"};
                return _dependents;
            }
        }

        public bool IsDuck
        {
            get
            {
                return false;
            }
        }

        public Type[] ExportedTypes
        {
            get
            {
                if (null == _exportedTypes)
                    _exportedTypes = Assembly.GetExportedTypes();
                return _exportedTypes;
            }
        }

        public bool Contains(Type type)
        {          
            foreach (Type item in ExportedTypes)
            {
                if (item == type)
                    return true;
            }

            return false;
        }

        public bool Contains(string className)
        {
            foreach (Type item in ExportedTypes)
            {
                if (item.Name.EndsWith(className, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        public bool ContractAndImplementation(Guid typeId, ref Type contract, ref Type implementation)
        {
            CreateTypesDictionary();
            foreach (var item in _types)
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

        public bool Implementation(Type contract, ref Type implementation)
        {
            CreateTypesDictionary();
            return _types.TryGetValue(contract, out implementation);
        }

        #endregion

        #region Methods

        private void CreateTypesDictionary()
        {
            if (null == _types)
            {
                _types = new Dictionary<Type, Type>();
                var contracts = ExportedTypes.Where(e => e.IsInterface
                                && e.Namespace == "VBIDEApi.Utils"
                                && null == e.GetCustomAttribute<SyntaxBypassAttribute>());
                foreach (var contract in contracts)
                {
                    var implementation = Assembly.GetType(contract.Namespace + ".Behind." + contract.Name, true);
                    _types.Add(contract, implementation);
                }

            }
        }

        #endregion
    }
    #pragma warning restore
}
