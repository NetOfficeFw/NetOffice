using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.EventContracts
{
    /// <summary>
    /// EStyle
    /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B06-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EStyle
	{
		/// <summary>
		/// StyleChanged
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// BeforeStyleDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// QueryCancelStyleDelete
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		/// <summary>
		/// StyleDeleteCanceled
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);
	}

}
