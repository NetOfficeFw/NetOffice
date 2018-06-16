using System;
using System.Collections.Generic;

namespace NetOffice.PowerPointApi.Tools.Expose
{
    /// <summary>
    /// OfficeApi Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.PowerPointApi";
        private Guid _componentId = new Guid("91493440-5A91-11CF-8700-00AA0060263B");
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
