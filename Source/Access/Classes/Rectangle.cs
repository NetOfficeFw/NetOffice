using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Rectangle_ClickEventHandler();
	public delegate void Rectangle_DblClickEventHandler(ref Int16 cancel);
	public delegate void Rectangle_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Rectangle_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Rectangle_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Rectangle 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836237.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._RectangleEvents), typeof(EventContracts.DispRectangleEvents))]
	[TypeId("3B06E949-E47C-11CD-8701-00AA003F0F07")]
    public interface Rectangle : _Rectangle, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197703.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Rectangle_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834418.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Rectangle_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845342.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Rectangle_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192713.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Rectangle_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834363.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Rectangle_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
