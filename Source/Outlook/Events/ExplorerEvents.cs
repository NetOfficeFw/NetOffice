using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// ExplorerEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006304F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorerEvents
	{
        /// <summary>
        /// Activate
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Activate();

        /// <summary>
        /// FolderSwitch
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void FolderSwitch();

        /// <summary>
        /// BeforeFolderSwitch
        /// </summary>
        /// <param name="newFolder"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newFolder", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel);

        /// <summary>
        /// ViewSwitch
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void ViewSwitch();

        /// <summary>
        /// BeforeViewSwitch
        /// </summary>
        /// <param name="newView"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel);

        /// <summary>
        /// Deactivate
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Deactivate();

        /// <summary>
        /// SelectionChange
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void SelectionChange();

        /// <summary>
        /// Close
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void Close();
	}
}
