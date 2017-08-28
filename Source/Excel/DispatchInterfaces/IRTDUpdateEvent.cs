using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	#pragma warning disable
	/// <summary>
	/// DispatchInterface IRTDUpdateEvent 
    /// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IRTDUpdateEvent
	{
		#region Methods

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(10)]
		void UpdateNotify();

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(12)]
		void Disconnect();

		#endregion

		#region Properties

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[DispId(11)]
		Int32 HeartbeatInterval{[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(11)] get;}

		#endregion
	}
}
	#pragma warning restore



