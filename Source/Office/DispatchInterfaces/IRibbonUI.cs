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
	/// DispatchInterface IRibbonUI SupportByVersionAttribute Office, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public interface IRibbonUI
	{
		#region Methods

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		void Invalidate();

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)]
		void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersionAttribute("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)]
		void InvalidateControlMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersionAttribute("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(4)]
		void ActivateTab([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersionAttribute("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(5)]
		void ActivateTabMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersionAttribute("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(6)]
		void ActivateTabQ([In, MarshalAs(UnmanagedType.BStr)]string ControlID, [In, MarshalAs(UnmanagedType.BStr)]string Namespace);

		#endregion

		#region Properties

		#endregion
	}
}
	#pragma warning restore
