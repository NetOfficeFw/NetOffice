using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVReplaceShapesEvent 
	/// SupportByVersion Visio, 15, 16
	/// </summary>
	[SupportByVersion("Visio", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0741-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.ReplaceShapesEvent))]
    public interface IVReplaceShapesEvent : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 15, 16), ProxyResult]
		object ReplacementMaster { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVSelection SelectionSource { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int32 ReplaceFlags { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		Int16 Stat { get; }

		#endregion

	}
}
