using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.EventContracts
{
    /// <summary>
    /// ControlEvents
    /// </summary>
	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("9A4BBF53-4E46-101B-8BBD-00AA003E3B29"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ControlEvents
	{
        /// <summary>
        /// Enter
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384830)]
		void Enter();

        /// <summary>
        /// Exit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384829)]
		void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// BeforeUpdate
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(NetOffice.MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384831)]
		void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

        /// <summary>
        /// AfterUpdate
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384832)]
		void AfterUpdate();
	}
}
