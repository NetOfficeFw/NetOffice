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
	/// DispatchInterface IRibbonUI SupportByVersionAttribute Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IRibbonUI
	{
		#region Methods

		[SupportByVersion("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		void Invalidate();

		[SupportByVersion("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)]
		void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)]
		void InvalidateControlMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(4)]
		void ActivateTab([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(5)]
		void ActivateTabMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(6)]
		void ActivateTabQ([In, MarshalAs(UnmanagedType.BStr)]string ControlID, [In, MarshalAs(UnmanagedType.BStr)]string Namespace);

		#endregion

		#region Properties

		#endregion
	}
}
	#pragma warning restore


