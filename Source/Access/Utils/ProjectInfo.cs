using System;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using NetOffice;

namespace AccessApi.Utils
{
	#pragma warning disable
    /// <summary>
    /// necessary factory info, used from NetOffice.Factory while Initialize()
    /// </summary>
    public class ProjectInfo : IFactoryInfo
    {
        #region Fields

        private string   _namespace     = "NetOffice.AccessApi";
        private Guid[]    _componentGuid = new Guid[]{new Guid("4AFFC9A0-5F99-101B-AF4E-00AA003F0F07")};
        private Assembly _assembly;
		private Type[]	 _exportedTypes;
		private string[] _dependents;
		
        #endregion

        #region Construction

        public ProjectInfo()
        {
            _assembly = Assembly.GetExecutingAssembly();
        }

        #endregion

        #region IFactoryInfo Members

		public bool Contains(string className)
		{
			if(null == _exportedTypes)
				_exportedTypes = Assembly.GetExportedTypes();
			
			foreach (Type item in _exportedTypes)
            {
				if (item.Name.EndsWith(className, StringComparison.InvariantCultureIgnoreCase))
					return true;
            }
				
			return false;			
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

        public string[] Dependencies
        {
            get
            {
				if(null == _dependents)
					_dependents = new string[]{"OfficeApi.dll","DAOApi.dll","VBIDEApi.dll","ADODBApi.dll","OWC10Api.dll"};
                return _dependents;
            }
        }
        
        #endregion
    }
    #pragma warning restore
}
