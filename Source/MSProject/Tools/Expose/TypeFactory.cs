using System;
using System.Collections.Generic;

namespace NetOffice.MSProjectApi.Tools.Expose
{
    /// <summary>
    /// MSProject Default Type Factory
    /// </summary>
    public class TypeFactory : NetOffice.Tools.Expose.Factory
    {
        private string _factoryNamespace = "NetOffice.MSProjectApi";
        private Guid _componentId = new Guid("A7107640-94DF-1068-855E-00DD01075445");
        private string[] _dependencies = new string[] { "OfficeApi.dll", "VBIDEApi.dll", "MSHTMLApi.dll" };

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
