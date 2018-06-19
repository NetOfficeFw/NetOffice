using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.EventContracts
{
    /// <summary>
    /// ApplicationEvents
    /// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00021240-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents
	{
		/// <summary>
		/// WindowActivate
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("wn", typeof(PublisherApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		/// <summary>
		/// WindowDeactivate
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("wn", typeof(PublisherApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		/// <summary>
		/// WindowPageChange
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("vw", typeof(PublisherApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw);

		/// <summary>
		/// Quit
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Quit();

		/// <summary>
		/// NewDocument
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentOpen
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// DocumentBeforeClose
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		/// <summary>
		/// MailMergeAfterMerge
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// MailMergeAfterRecordMerge
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// MailMergeBeforeMerge
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("startRecord", SinkArgumentType.Int32)]
        [SinkArgument("endRecord", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		/// <summary>
		/// MailMergeBeforeRecordMerge
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		/// <summary>
		/// MailMergeDataSourceLoad
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// MailMergeWizardSendToCustom
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// MailMergeWizardStateChange
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("fromState", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object fromState);

		/// <summary>
		/// MailMergeDataSourceValidate
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		/// <summary>
		/// MailMergeInsertBarcode
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("okToInsert", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeInsertBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object okToInsert);

		/// <summary>
		/// MailMergeRecipientListClose
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeRecipientListClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// MailMergeGenerateBarcode
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("bstrString", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeGenerateBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object bstrString);

		/// <summary>
		/// MailMergeWizardFollowUpCustom
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeWizardFollowUpCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// BeforePrint
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		/// <summary>
		/// AfterPrint
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		/// <summary>
		/// ShowCatalogUI
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ShowCatalogUI();

		/// <summary>
		/// HideCatalogUI
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void HideCatalogUI();
	}
}
