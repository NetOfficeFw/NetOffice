using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CustomControl_UpdatedEventHandler(ref Int16 code);
	public delegate void CustomControl_EnterEventHandler();
	public delegate void CustomControl_ExitEventHandler(ref Int16 cancel);
	public delegate void CustomControl_GotFocusEventHandler();
	public delegate void CustomControl_LostFocusEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass CustomControl 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821191.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._CustomControlEvents), typeof(EventContracts.DispCustomControlEvents))]
	[TypeId("3B06E967-E47C-11CD-8701-00AA003F0F07")]
    public interface CustomControl : _CustomControl, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193546.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CustomControl_UpdatedEventHandler UpdatedEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836867.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CustomControl_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192738.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CustomControl_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822815.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CustomControl_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844932.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event CustomControl_LostFocusEventHandler LostFocusEvent;

		#endregion
	}
}
