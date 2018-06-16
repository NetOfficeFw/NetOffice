using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.EventContracts
{
    /// <summary>
    /// _DataSourceControlEvent
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F5B39A9B-1480-11D3-8549-00C04FAC67D7"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DataSourceControlEvent
	{
        /// <summary>
        /// Current
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(624)]
		void Current([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeExpand
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(626)]
		void BeforeExpand([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeCollapse
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(627)]
		void BeforeCollapse([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeFirstPage
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(628)]
		void BeforeFirstPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforePreviousPage
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(629)]
		void BeforePreviousPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeNextPage
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(630)]
		void BeforeNextPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeLastPage
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(631)]
		void BeforeLastPage([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// DataError
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(632)]
		void DataError([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// DataPageComplete
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(633)]
		void DataPageComplete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeInitialBind
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(634)]
		void BeforeInitialBind([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// RecordsetSaveProgress
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(635)]
		void RecordsetSaveProgress([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// AfterDelete
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(636)]
		void AfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// AfterInsert
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(637)]
		void AfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// AfterUpdate
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(638)]
		void AfterUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeDelete
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(639)]
		void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeInsert
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(640)]
		void BeforeInsert([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeOverwrite
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(641)]
		void BeforeOverwrite([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// BeforeUpdate
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(642)]
		void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// Dirty
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(643)]
		void Dirty([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// RecordExit
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(644)]
		void RecordExit([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// Undo
        /// </summary>
        /// <param name="dSCEventInfo"></param>
        [SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(647)]
		void Undo([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);

        /// <summary>
        /// Focus
        /// </summary>
        /// <param name="dSCEventInfo"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("dSCEventInfo", typeof(OWC10Api.DSCEventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(648)]
		void Focus([In, MarshalAs(UnmanagedType.IDispatch)] object dSCEventInfo);
	}
}
