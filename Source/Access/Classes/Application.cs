using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CoreServices;

namespace NetOffice.AccessApi
{
    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.AccessApi.Behind.Application
    /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.AccessApi.Behind.Application
    {
        private string _defaultProgId = "Access.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Access
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Access based on given id.
        /// This can be used to target a specific version of Microsoft Access.
        /// Example usage:
        /// "Access.Excel.12" to target Access 2007
        /// "Access.Excel.14" to target Access 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Access
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Access
        /// </summary>
        /// <param name="mode">indicates where is the call coming from</param>
        public ApplicationClass(NetOffice.Callers.InteropCompatibilityClassCreateMode mode)
        {
            if (mode == NetOffice.Callers.InteropCompatibilityClassCreateMode.Direct)
            {
                ICOMObjectInitialize init = (ICOMObjectInitialize)this;
                init.InitializeCOMObject(_defaultProgId);
            }
        }
    }

    /// <summary>
    /// CoClass Application
    /// SupportByVersion Access, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821758.aspx </remarks>
    [SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Access.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
	[TypeId("73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9")]
    public interface Application : _Application, ICloneable<Application>, ICOMObjectProxyService
	{

	}
}
