using System;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using NetOffice;

namespace MSComctlLibApi.Utils
{
	#pragma warning disable
    /// <summary>
    /// necessary factory info, used from NetOffice.Factory while Initialize()
    /// </summary>
    public class ProjectInfo : IFactoryInfo
    {
        #region Fields

        private string   _namespace     = "NetOffice.MSComctlLibApi";
        private Guid[]    _componentGuid = new Guid[]{new Guid("831FDD16-0C5C-11D2-A9FC-0000F8754DA1")};
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
					_dependents = new string[0];
                return _dependents;
            }
        }
        
        #endregion
    }
    #pragma warning restore
}
