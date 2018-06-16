using System;
using System.Collections.Generic;

namespace NetOffice.MSDATASRCApi.Tools.Expose
{
    /// <summary>
    /// MSDATASRC Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.MSDATASRCApi";
        private Guid _componentId = new Guid("7C0FFAB0-CD84-11D0-949A-00A0C91110ED");
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
