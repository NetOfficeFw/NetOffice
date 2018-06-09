using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SmartDocument 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863963.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000C0377-0000-0000-C000-000000000046")]
    public interface SmartDocument : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864983.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string SolutionID { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865469.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string SolutionURL { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865250.aspx </remarks>
        /// <param name="considerAllSchemas">optional bool ConsiderAllSchemas = false</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void PickSolution(object considerAllSchemas);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865250.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void PickSolution();

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864173.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void RefreshPane();

        #endregion
    }
}
