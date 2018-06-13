using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void NameSpace_OptionsPagesAddEventHandler(NetOffice.OutlookApi.PropertyPages pages, NetOffice.OutlookApi.MAPIFolder folder);
	public delegate void NameSpace_AutoDiscoverCompleteEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass NameSpace 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869848.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.NameSpaceEvents))]
	[TypeId("0006308B-0000-0000-C000-000000000046")]
    public interface NameSpace : _NameSpace, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863940.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event NameSpace_OptionsPagesAddEventHandler OptionsPagesAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868722.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event NameSpace_AutoDiscoverCompleteEventHandler AutoDiscoverCompleteEvent;

        #endregion
    }
}
