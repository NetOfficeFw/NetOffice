using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// SyncObjectEvents
    /// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063085-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface SyncObjectEvents
	{
        /// <summary>
        /// SyncStart
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void SyncStart();

        /// <summary>
        /// Progress
        /// </summary>
        /// <param name="state"></param>
        /// <param name="description"></param>
        /// <param name="value"></param>
        /// <param name="max"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("state", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlSyncState))]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("value", SinkArgumentType.Int32)]
        [SinkArgument("max", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void Progress([In] object state, [In] object description, [In] object value, [In] object max);

        /// <summary>
        /// OnError
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("code", SinkArgumentType.Int32)]
        [SinkArgument("description", SinkArgumentType.String)]       
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void OnError([In] object code, [In] object description);

        /// <summary>
        /// SyncEnd
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void SyncEnd();
	}
}
