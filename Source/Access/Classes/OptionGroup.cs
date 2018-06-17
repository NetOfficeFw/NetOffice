using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OptionGroup_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void OptionGroup_AfterUpdateEventHandler();
	public delegate void OptionGroup_EnterEventHandler();
	public delegate void OptionGroup_ExitEventHandler(ref Int16 cancel);
	public delegate void OptionGroup_ClickEventHandler();
	public delegate void OptionGroup_DblClickEventHandler(ref Int16 cancel);
	public delegate void OptionGroup_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void OptionGroup_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void OptionGroup_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OptionGroup 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821461.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._OptionGroupEvents), typeof(EventContracts.DispOptionGroupEvents))]
	[TypeId("3B06E955-E47C-11CD-8701-00AA003F0F07")]
    public interface OptionGroup : _OptionGroup, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821100.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836238.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821475.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192101.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196181.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193768.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836672.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195835.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845867.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event OptionGroup_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
