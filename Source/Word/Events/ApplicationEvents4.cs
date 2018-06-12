using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.EventContracts
{
    /// <summary>
    /// ApplicationEvents4
    /// </summary>
    [SupportByVersion("Word", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00020A01-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents4
	{
        /// <summary>
        /// Startup
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

        /// <summary>
        /// Quit
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

        /// <summary>
        /// DocumentChange
        /// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

        /// <summary>
        /// DocumentOpen
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// DocumentBeforeClose
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

        /// <summary>
        /// DocumentBeforePrint
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

        /// <summary>
        /// DocumentBeforeSave
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("saveAsUI", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

        /// <summary>
        /// NewDocument
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// WindowActivate
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowDeactivate
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// WindowSelectionChange
        /// </summary>
        /// <param name="sel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

        /// <summary>
        /// WindowBeforeRightClick
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

        /// <summary>
        /// WindowBeforeDoubleClick
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

        /// <summary>
        /// EPostagePropertyDialog
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// EPostageInsert
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// MailMergeAfterMerge
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="docResult"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("docResult", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult);

        /// <summary>
        /// MailMergeAfterRecordMerge
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// MailMergeBeforeMerge
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="startRecord"></param>
        /// <param name="endRecord"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("startRecord", SinkArgumentType.Int32)]
        [SinkArgument("endRecord", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

        /// <summary>
        /// MailMergeBeforeRecordMerge
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

        /// <summary>
        /// MailMergeDataSourceLoad
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// MailMergeDataSourceValidate
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="handled"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

        /// <summary>
        /// MailMergeWizardSendToCustom
        /// </summary>
        /// <param name="doc"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

        /// <summary>
        /// MailMergeWizardStateChange
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="fromState"></param>
        /// <param name="toState"></param>
        /// <param name="handled"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("fromState", SinkArgumentType.Int32)]
        [SinkArgument("toState", SinkArgumentType.Int32)]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled);

        /// <summary>
        /// WindowSize
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("wn", typeof(WordApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

        /// <summary>
        /// XMLSelectionChange
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="oldXMLNode"></param>
        /// <param name="newXMLNode"></param>
        /// <param name="reason"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("sel", typeof(WordApi.Selection))]
        [SinkArgument("oldXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("newXMLNode", typeof(WordApi.XMLNode))]
        [SinkArgument("reason", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void XMLSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] [Out] ref object reason);

        /// <summary>
        /// XMLValidationError
        /// </summary>
        /// <param name="xMLNode"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("xMLNode", typeof(WordApi.XMLNode))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void XMLValidationError([In, MarshalAs(UnmanagedType.IDispatch)] object xMLNode);

        /// <summary>
        /// DocumentSync
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="syncEventType"></param>
        [SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("syncEventType", SinkArgumentType.Enum, typeof(OfficeApi.Enums.MsoSyncEventType))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void DocumentSync([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object syncEventType);

        /// <summary>
        /// EPostageInsertEx
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cpDeliveryAddrStart"></param>
        /// <param name="cpDeliveryAddrEnd"></param>
        /// <param name="cpReturnAddrStart"></param>
        /// <param name="cpReturnAddrEnd"></param>
        /// <param name="xaWidth"></param>
        /// <param name="yaHeight"></param>
        /// <param name="bstrPrinterName"></param>
        /// <param name="bstrPaperFeed"></param>
        /// <param name="fPrint"></param>
        /// <param name="fCancel"></param>
		[SupportByVersion("Word", 11,12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("cpDeliveryAddrStart", SinkArgumentType.Int32)]
        [SinkArgument("cpDeliveryAddrEnd", SinkArgumentType.Int32)]
        [SinkArgument("cpReturnAddrStart", SinkArgumentType.Int32)]
        [SinkArgument("cpReturnAddrEnd", SinkArgumentType.Int32)]
        [SinkArgument("xaWidth", SinkArgumentType.Int32)]
        [SinkArgument("yaHeight", SinkArgumentType.Int32)]
        [SinkArgument("bstrPrinterName", SinkArgumentType.String)]
        [SinkArgument("bstrPaperFeed", SinkArgumentType.String)]
        [SinkArgument("fPrint", SinkArgumentType.Bool)]
        [SinkArgument("fCancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(29)]
		void EPostageInsertEx([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object cpDeliveryAddrStart, [In] object cpDeliveryAddrEnd, [In] object cpReturnAddrStart, [In] object cpReturnAddrEnd, [In] object xaWidth, [In] object yaHeight, [In] object bstrPrinterName, [In] object bstrPaperFeed, [In] object fPrint, [In] [Out] ref object fCancel);

        /// <summary>
        /// MailMergeDataSourceValidate2
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="handled"></param>
		[SupportByVersion("Word", 12,14,15,16)]
        [SinkArgument("doc", typeof(WordApi.Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void MailMergeDataSourceValidate2([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

        /// <summary>
        /// ProtectedViewWindowOpen
        /// </summary>
        /// <param name="pvWindow"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

        /// <summary>
        /// ProtectedViewWindowBeforeEdit
        /// </summary>
        /// <param name="pvWindow"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowBeforeClose
        /// </summary>
        /// <param name="pvWindow"></param>
        /// <param name="closeReason"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [SinkArgument("newCloseReason", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] object closeReason, [In] [Out] ref object cancel);

        /// <summary>
        /// ProtectedViewWindowSize
        /// </summary>
        /// <param name="pvWindow"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void ProtectedViewWindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

        /// <summary>
        /// ProtectedViewWindowActivate
        /// </summary>
        /// <param name="pvWindow"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

        /// <summary>
        /// ProtectedViewWindowDeactivate
        /// </summary>
        /// <param name="pvWindow"></param>
		[SupportByVersion("Word", 14,15,16)]
        [SinkArgument("pvWindow", typeof(WordApi.ProtectedViewWindow))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);
	}
}
