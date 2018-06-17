using System;
using System.Collections.Generic;

namespace NetOffice.AccessApi.Tools.Expose
{
    /// <summary>
    /// Office Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.AccessApi";
        private Guid _componentId = new Guid("4AFFC9A0-5F99-101B-AF4E-00AA003F0F07");
        private string[] _dependencies = new string[] { "OfficeApi.dll", "DAOApi.dll", "VBIDEApi.dll", "ADODBApi.dll", "OWC10Api.dll" };

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
