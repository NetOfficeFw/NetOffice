using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// FormRegionEvents
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006305B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface FormRegionEvents
	{
        /// <summary>
        /// Expanded
        /// </summary>
        /// <param name="expand"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("expand", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64312)]
		void Expanded([In] object expand);

        /// <summary>
        /// Close
        /// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Close();
	}
}
