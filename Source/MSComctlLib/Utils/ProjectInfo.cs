﻿using System;
using System.Reflection;
using NetOffice;
using NetOffice.Attributes;

namespace MSComctlLibApi.Utils
{
    #pragma warning disable
    /// <summary>
    /// Necessary factory info, used from NetOffice.Core while Initialize()
    /// </summary>
    public class ProjectInfo : IFactoryInfo
    {
        #region Fields

        private string    _name;
        private string    _namespace     = "NetOffice.MSComctlLibApi";
        private Guid[]    _componentGuid = new Guid[]{new Guid("831FDD16-0C5C-11D2-A9FC-0000F8754DA1")};
        private Assembly  _assembly;
        private NetOfficeAssemblyAttribute _assemblyAttribute;
        private Type[]	  _exportedTypes;
        private string[]  _dependents;

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

        public Guid[] ComponentGuid
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
					_dependents = new string[0];
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

        public bool Contains(Type type)
        {
            if (null == _exportedTypes)
                _exportedTypes = Assembly.GetExportedTypes();

            foreach (Type item in _exportedTypes)
            {
                if (item == type)
                    return true;
            }

            return false;
        }

        public bool Contains(string className)
        {
            if (null == _exportedTypes)
                _exportedTypes = Assembly.GetExportedTypes();

            foreach (Type item in _exportedTypes)
            {
                if (item.Name.EndsWith(className, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        #endregion
    }
    #pragma warning restore
}
