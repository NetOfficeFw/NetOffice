using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ApplicationEvents_10
    /// </summary>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006300E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents_10
	{
        /// <summary>
        /// ItemSend
        /// </summary>
        /// <param name="item"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

        /// <summary>
        /// NewMail
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void NewMail();

        /// <summary>
        /// Reminder
        /// </summary>
        /// <param name="item"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        /// <summary>
        /// OptionsPagesAdd
        /// </summary>
        /// <param name="pages"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("pages", typeof(OutlookApi.PropertyPages))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages);

        /// <summary>
        /// Startup
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Startup();

        /// <summary>
        /// Quit
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void Quit();

        /// <summary>
        /// AdvancedSearchComplete
        /// </summary>
        /// <param name="searchObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64106)]
		void AdvancedSearchComplete([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

        /// <summary>
        /// AdvancedSearchStopped
        /// </summary>
        /// <param name="searchObject"></param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("searchObject", typeof(OutlookApi.Search))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64107)]
		void AdvancedSearchStopped([In, MarshalAs(UnmanagedType.IDispatch)] object searchObject);

        /// <summary>
        /// MAPILogonComplete
        /// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64144)]
		void MAPILogonComplete();
	}

}
