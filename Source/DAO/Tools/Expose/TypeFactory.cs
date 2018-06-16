using System;
using System.Collections.Generic;

namespace NetOffice.DAOApi.Tools.Expose
{
    /// <summary>
    /// DAOApi Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.DAOApi";
        private Guid _componentId = new Guid("00025E01-0000-0000-C000-000000000046"); // new Guid("4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28") 
        private string[] _dependencies = new string[0];

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
