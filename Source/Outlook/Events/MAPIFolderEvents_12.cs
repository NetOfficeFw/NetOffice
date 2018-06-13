using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// MAPIFolderEvents_12
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F7-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MAPIFolderEvents_12
	{
        /// <summary>
        /// BeforeFolderMove
        /// </summary>
        /// <param name="moveTo"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("moveTo", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64424)]
		void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);

        /// <summary>
        /// BeforeItemMove
        /// </summary>
        /// <param name="item"></param>
        /// <param name="moveTo"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("moveTo", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64425)]
		void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);
	}
}
