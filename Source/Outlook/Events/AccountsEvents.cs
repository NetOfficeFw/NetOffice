using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// AccountsEvents
    /// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063105-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface AccountsEvents
	{
        /// <summary>
        /// AutoDiscoverComplete
        /// </summary>
        /// <param name="account"></param>
		[SupportByVersion("Outlook", 14,15,16)]
        [SinkArgument("account", typeof(OutlookApi.Account))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64620)]
		void AutoDiscoverComplete([In, MarshalAs(UnmanagedType.IDispatch)] object account);
	}
}
