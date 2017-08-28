using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#pragma warning disable
	/// <summary>
	/// DispatchInterface IRibbonControl SupportByVersionAttribute Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IRibbonControl
	{
		#region Methods

		#endregion

		#region Properties

		[SupportByVersion("Office", 12,14,15,16)]
		[DispId(1)]
		string Id{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)] get;}

		[SupportByVersion("Office", 12,14,15,16)]
		[DispId(2)]
		object Context{[return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)] get;}

		[SupportByVersion("Office", 12,14,15,16)]
		[DispId(3)]
		string Tag{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)] get;}

		#endregion
	}
}
	#pragma warning restore


