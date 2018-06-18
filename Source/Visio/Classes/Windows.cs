using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Windows_WindowOpenedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Windows_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void Windows_MouseDownEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_MouseMoveEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_MouseUpEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Windows_KeyDownEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void Windows_KeyPressEventHandler(Int32 keyAscii, ref bool CancelDefault);
	public delegate void Windows_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Windows 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769453(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EWindows))]
	[TypeId("000D0A0B-0000-0000-C000-000000000046")]
    public interface Windows : IVWindows, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765893(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_WindowOpenedEventHandler WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765788(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_SelectionChangedEventHandler SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769123(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_BeforeWindowClosedEventHandler BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766087(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_WindowActivatedEventHandler WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768566(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768773(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768369(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_WindowTurnedToPageEventHandler WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765071(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_WindowChangedEventHandler WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765224(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_ViewChangedEventHandler ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768016(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766052(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_WindowCloseCanceledEventHandler WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766300(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768824(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766046(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765457(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766065(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766041(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765374(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Windows_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
