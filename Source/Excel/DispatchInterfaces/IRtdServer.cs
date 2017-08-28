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
	/// DispatchInterface IRtdServer SupportByVersionAttribute Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IRtdServer
	{
		#region Methods

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(10)]
		Int32 ServerStart([In, MarshalAs(UnmanagedType.Interface)] object CallbackObject);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[return: MarshalAs(UnmanagedType.Struct)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(11)]
		object ConnectData([In]Int32 TopicID, [In] object Strings, [In]bool GetNewValues);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(12)]
		object RefreshData([In]Int32 TopicCount);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(13)]
		void DisconnectData([In]Int32 TopicID);

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(14)]
		Int32 Heartbeat();

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(15)]
		void ServerTerminate();

		#endregion

		#region Properties

		#endregion
	}
}
	#pragma warning restore



