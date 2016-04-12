using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ExcelApi
{
	#pragma warning disable
	///<summary>
	/// DispatchInterface IRTDUpdateEvent SupportByVersionAttribute Excel, 10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
	[ComImport, ComVisible(true), Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), TypeLibType((short) 4160)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public interface IRTDUpdateEvent
	{
		#region Methods

		[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(10)]
		void UpdateNotify();

		[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(12)]
		void Disconnect();

		#endregion

		#region Properties

		[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		[DispId(11)]
		Int32 HeartbeatInterval{[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(11)] get;}

		#endregion
	}
}
	#pragma warning restore
