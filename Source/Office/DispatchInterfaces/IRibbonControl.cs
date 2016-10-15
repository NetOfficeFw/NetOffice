using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	#pragma warning disable
	///<summary>
	/// DispatchInterface IRibbonControl SupportByVersionAttribute Office, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public interface IRibbonControl
	{
		#region Methods

		#endregion

		#region Properties

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[DispId(1)]
		string Id{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)] get;}

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[DispId(2)]
		object Context{[return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)] get;}

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[DispId(3)]
		string Tag{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)] get;}

		#endregion
	}
}
	#pragma warning restore
