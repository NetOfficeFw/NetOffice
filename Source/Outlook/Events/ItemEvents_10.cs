using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ItemEvents_10
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006302B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ItemEvents_10
	{
        /// <summary>
        /// Open
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void Open([In] [Out] ref object cancel);

        /// <summary>
        /// CustomAction
        /// </summary>
        /// <param name="action"></param>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("action", SinkArgumentType.UnknownProxy)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void CustomAction([In, MarshalAs(UnmanagedType.IDispatch)] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

        /// <summary>
        /// CustomPropertyChange
        /// </summary>
        /// <param name="name"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("name", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void CustomPropertyChange([In] object name);

        /// <summary>
        /// Forward
        /// </summary>
        /// <param name="forward"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("forward", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62568)]
		void Forward([In, MarshalAs(UnmanagedType.IDispatch)] object forward, [In] [Out] ref object cancel);

        /// <summary>
        /// Close
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Close([In] [Out] ref object cancel);

        /// <summary>
        /// PropertyChange
        /// </summary>
        /// <param name="name"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("name", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61449)]
		void PropertyChange([In] object name);

        /// <summary>
        /// Read
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Read();

        /// <summary>
        /// Reply
        /// </summary>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62566)]
		void Reply([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

        /// <summary>
        /// ReplyAll
        /// </summary>
        /// <param name="response"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62567)]
		void ReplyAll([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

        /// <summary>
        /// Send
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void Send([In] [Out] ref object cancel);

        /// <summary>
        /// Write
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void Write([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeCheckNames
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61450)]
		void BeforeCheckNames([In] [Out] ref object cancel);

        /// <summary>
        /// AttachmentAdd
        /// </summary>
        /// <param name="attachment"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61451)]
		void AttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

        /// <summary>
        /// AttachmentRead
        /// </summary>
        /// <param name="attachment"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61452)]
		void AttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

        /// <summary>
        /// BeforeAttachmentSave
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61453)]
		void BeforeAttachmentSave([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeDelete
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64117)]
		void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

        /// <summary>
        /// AttachmentRemove
        /// </summary>
        /// <param name="attachment"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64430)]
		void AttachmentRemove([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

        /// <summary>
        /// BeforeAttachmentAdd
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64432)]
		void BeforeAttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeAttachmentPreview
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64431)]
		void BeforeAttachmentPreview([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeAttachmentRead
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64427)]
		void BeforeAttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeAttachmentWriteToTempFile
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64434)]
		void BeforeAttachmentWriteToTempFile([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

        /// <summary>
        /// Unload
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64429)]
		void Unload();

        /// <summary>
        /// BeforeAutoSave
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64514)]
		void BeforeAutoSave([In] [Out] ref object cancel);

        /// <summary>
        /// BeforeRead
        /// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64652)]
		void BeforeRead();

        /// <summary>
        /// AfterWrite
        /// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64653)]
		void AfterWrite();

        /// <summary>
        /// ReadComplete
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 15, 16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64655)]
		void ReadComplete([In] [Out] ref object cancel);
	}
}
