using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Section_ClickEventHandler();
	public delegate void Section_DblClickEventHandler(ref Int16 cancel);
	public delegate void Section_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Section_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Section_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Section_PaintEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Section 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198334.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._SectionEvents), typeof(EventContracts.DispSectionEvents))]
	[TypeId("BC9E4355-F037-11CD-8701-00AA003F0F07")]
    public interface Section : _Section, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835739.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Section_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195125.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Section_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835713.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Section_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194501.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Section_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837223.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Section_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836875.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		event Section_PaintEventHandler PaintEvent;

		#endregion
	}
}
