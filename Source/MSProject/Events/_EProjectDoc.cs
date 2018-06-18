using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.EventContracts
{
    /// <summary>
    /// _EProjectDoc
    /// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F81DD3C0-5089-11CF-A49D-00AA00574C74"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectDoc
	{
        /// <summary>
        /// Open
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// BeforeClose
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// BeforeSave
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// BeforePrint
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// Calculate
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// Change
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// Activate
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// Deactivate
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);
	}
}
