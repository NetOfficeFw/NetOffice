using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void FormRegion_ExpandedEventHandler(bool expand);
	public delegate void FormRegion_CloseEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass FormRegion 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863634.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.FormRegionEvents))]
	[TypeId("0006315A-0000-0000-C000-000000000046")]
    public interface FormRegion : _FormRegion, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868186.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event FormRegion_ExpandedEventHandler ExpandedEvent;

		/// <summary>
		/// SupportByVersion Outlook 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860943.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		event FormRegion_CloseEventHandler CloseEvent;

        #endregion
    }
}
