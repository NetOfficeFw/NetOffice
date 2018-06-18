using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Styles_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Styles_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Styles_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Styles_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Styles_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle style);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Styles 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769402(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EStyles))]
	[TypeId("000D0A01-0000-0000-C000-000000000046")]
    public interface Styles : IVStyles, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768234(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Styles_StyleAddedEventHandler StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766489(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Styles_StyleChangedEventHandler StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768770(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Styles_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765172(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Styles_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767143(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Styles_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent;

		#endregion
	}
}
