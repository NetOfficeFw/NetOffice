using System;
using System.Collections.Generic;

namespace NetOffice.WordApi.Tools.Expose
{
    /// <summary>
    /// WordApi Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.WordApi";
        private Guid _componentId = new Guid("00020905-0000-0000-C000-000000000046");
        private string[] _dependencies = new string[] { "OfficeApi.dll", "VBIDEApi.dll" };

        /// <summary>
        /// Default namespace of the factory
        /// </summary>
        public override string FactoryNamespace
        {
            get
            {
                return _factoryNamespace;
            }
        }

        /// <summary>
        /// Guid of the COM component which represents the NetOfficeApi assembly
        /// </summary>
        public override Guid ComponentID
        {
            get
            {
                return _componentId;
            }
        }

        /// <summary>
        /// Returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        public override string[] Dependencies
        {
            get
            {
                return _dependencies;
            }
        }
    }
}
