using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.EventContracts
{
    /// <summary>
    /// DocumentEvents2
    /// </summary>
    [SupportByVersion("Word", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00020A02-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocumentEvents2
	{
        /// <summary>
        /// New
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void New();

        /// <summary>
        /// Open
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Open();

        /// <summary>
        /// Close
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Close();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="syncEventType"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Sync([In] object syncEventType);

        /// <summary>
        /// XMLAfterInsert
        /// </summary>
        /// <param name="newXMLNode"></param>
        /// <param name="inUndoRedo"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("newXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo);

        /// <summary>
        /// XMLBeforeDelete
        /// </summary>
        /// <param name="deletedRange"></param>
        /// <param name="oldXMLNode"></param>
        /// <param name="inUndoRedo"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("deletedRange", typeof(WordApi.Range))]
        [SinkArgument("oldXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo);

        /// <summary>
        /// ContentControlAfterAdd
        /// </summary>
        /// <param name="newContentControl"></param>
        /// <param name="inUndoRedo"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("newContentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo);

        /// <summary>
        /// ContentControlBeforeDelete
        /// </summary>
        /// <param name="oldContentControl"></param>
        /// <param name="inUndoRedo"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("oldContentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo);

        /// <summary>
        /// ContentControlOnExit
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel);

        /// <summary>
        /// ContentControlOnEnter
        /// </summary>
        /// <param name="contentControl"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl);

        /// <summary>
        /// ContentControlBeforeStoreUpdate
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="content"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("content", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

        /// <summary>
        /// ContentControlBeforeContentUpdate
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="content"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("contentControl", typeof(WordApi.ContentControl))]
        [SinkArgument("content", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content);

        /// <summary>
        /// BuildingBlockInsert
        /// </summary>
        /// <param name="range"></param>
        /// <param name="name"></param>
        /// <param name="category"></param>
        /// <param name="blockType"></param>
        /// <param name="template"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("range", typeof(WordApi.Range))]
        [SinkArgument("name", SinkArgumentType.String)]
        [SinkArgument("category", SinkArgumentType.String)]
        [SinkArgument("blockType", SinkArgumentType.String)]
        [SinkArgument("template", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template);
	}
}
