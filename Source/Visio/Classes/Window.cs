using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Window_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow window);
	public delegate void Window_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap msg);
	public delegate void Window_MouseDownEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_MouseMoveEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_MouseUpEventHandler(Int32 button, Int32 keyButtonState, Double x, Double y, ref bool cancelDefault);
	public delegate void Window_KeyDownEventHandler(Int32 KeyCode, Int32 keyButtonState, ref bool cancelDefault);
	public delegate void Window_KeyPressEventHandler(Int32 keyAscii, ref bool cancelDefault);
	public delegate void Window_KeyUpEventHandler(Int32 keyCode, Int32 keyButtonState, ref bool cancelDefault);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Window 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769449(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EWindow))]
	[TypeId("000D0A0C-0000-0000-C000-000000000046")]
    public interface Window : IVWindow, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766358(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_SelectionChangedEventHandler SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766158(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_BeforeWindowClosedEventHandler BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767395(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_WindowActivatedEventHandler WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766155(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767156(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768977(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_WindowTurnedToPageEventHandler WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768918(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_WindowChangedEventHandler WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767710(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_ViewChangedEventHandler ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766114(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768122(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_WindowCloseCanceledEventHandler WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767265(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767550(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_MouseDownEventHandler MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767508(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_MouseMoveEventHandler MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_MouseUpEventHandler MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766901(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_KeyDownEventHandler KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767366(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_KeyPressEventHandler KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767880(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Window_KeyUpEventHandler KeyUpEvent;

		#endregion
	}
}
