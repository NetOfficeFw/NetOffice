using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Style_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Style_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Style_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle style);
	public delegate void Style_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle style);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Style 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769398(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EStyle))]
	[TypeId("000D0A02-0000-0000-C000-000000000046")]
    public interface Style : IVStyle, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768814(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Style_StyleChangedEventHandler StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765885(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Style_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766049(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Style_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765208(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Style_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent;

		#endregion
	}
}
