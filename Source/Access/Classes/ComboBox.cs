using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void ComboBox_BeforeUpdateEventHandler(ref Int16 cancel);
	public delegate void ComboBox_AfterUpdateEventHandler();
	public delegate void ComboBox_ChangeEventHandler();
	public delegate void ComboBox_NotInListEventHandler(ref string newData, ref Int16 response);
	public delegate void ComboBox_EnterEventHandler();
	public delegate void ComboBox_ExitEventHandler(ref Int16 cancel);
	public delegate void ComboBox_GotFocusEventHandler();
	public delegate void ComboBox_LostFocusEventHandler();
	public delegate void ComboBox_ClickEventHandler();
	public delegate void ComboBox_DblClickEventHandler(ref Int16 cancel);
	public delegate void ComboBox_MouseDownEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ComboBox_MouseMoveEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ComboBox_MouseUpEventHandler(ref Int16 button, ref Int16 shift, ref Single x, ref Single y);
	public delegate void ComboBox_KeyDownEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void ComboBox_KeyPressEventHandler(ref Int16 keyAscii);
	public delegate void ComboBox_KeyUpEventHandler(ref Int16 keyCode, ref Int16 shift);
	public delegate void ComboBox_DirtyEventHandler(ref Int16 cancel);
	public delegate void ComboBox_UndoEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass ComboBox
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845773.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._ComboBoxEvents), typeof(EventContracts.DispComboBoxEvents))]
	[TypeId("3B06E95B-E47C-11CD-8701-00AA003F0F07")]
    public interface ComboBox : _Combobox, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193544.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_BeforeUpdateEventHandler BeforeUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197081.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_AfterUpdateEventHandler AfterUpdateEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836326.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845736.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_NotInListEventHandler NotInListEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822059.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193235.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_ExitEventHandler ExitEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196346.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_GotFocusEventHandler GotFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835708.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_LostFocusEventHandler LostFocusEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196406.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_ClickEventHandler ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196034.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_DblClickEventHandler DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192689.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195865.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192866.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197678.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196758.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821477.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event ComboBox_KeyUpEventHandler KeyUpEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845487.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event ComboBox_DirtyEventHandler DirtyEvent;

		/// <summary>
		/// SupportByVersion Access 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834733.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		event ComboBox_UndoEventHandler UndoEvent;

		#endregion
	}
}
