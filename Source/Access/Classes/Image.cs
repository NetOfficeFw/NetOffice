using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Image_ClickEventHandler();
	public delegate void Image_DblClickEventHandler(ref Int16 cancel);
	public delegate void Image_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Image_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void Image_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Image 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845723.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ImageEvents), typeof(EventContracts.DispImageEvents))]
	[TypeId("3B06E94D-E47C-11CD-8701-00AA003F0F07")]
    public interface Image : _Image, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845722.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Image_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194808.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Image_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844812.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Image_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195115.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Image_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192041.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event Image_MouseUpEventHandler MouseUpEvent;

		#endregion
	}
}
