using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    #region Delegates

    #pragma warning disable
    public delegate void QueryTable_BeforeRefreshEventHandler(ref bool cancel);
    public delegate void QueryTable_AfterRefreshEventHandler(bool success);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass QueryTable
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198271.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.RefreshEvents))]
	[TypeId("59191DA1-EA47-11CE-A51F-00AA0061507F")]
    public interface QueryTable : _QueryTable, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823150.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event QueryTable_BeforeRefreshEventHandler BeforeRefreshEvent;

        /// <summary>
        /// SupportByVersion Excel 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835922.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event QueryTable_AfterRefreshEventHandler AfterRefreshEvent;

        #endregion
    }
}
