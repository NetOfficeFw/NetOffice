using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CommandBars_OnUpdateEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CommandBars
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860339.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.OfficeApi.EventContracts._CommandBarsEvents))]
	[TypeId("55F88893-7708-11D1-ACEB-006008961DA5")]
    public interface CommandBars : _CommandBars, IEventBinding
    {
        #region Events

        /// <summary>
		/// SupportByVersion Office 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861536.aspx </remarks>
		[SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        event CommandBars_OnUpdateEventHandler OnUpdateEvent;

        #endregion
    }
}
