using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void References_ItemAddedEventHandler(NetOffice.AccessApi.Reference reference);
	public delegate void References_ItemRemovedEventHandler(NetOffice.AccessApi.Reference reference);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass References 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821489.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._References_Events))]
	[TypeId("EB106214-9C89-11CF-A2B3-00A0C90542FF")]
    public interface References : _References, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823174.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event References_ItemAddedEventHandler ItemAddedEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845638.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event References_ItemRemovedEventHandler ItemRemovedEvent;

		#endregion
	}
}
