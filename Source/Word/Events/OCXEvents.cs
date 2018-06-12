using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.EventContracts
{
    /// <summary>
    /// OCXEvents
    /// </summary>
    [SupportByVersion("Word", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000209F3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OCXEvents
	{
        /// <summary>
        /// GotFocus
        /// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147417888)]
		void GotFocus();

        /// <summary>
        /// LostFocus
        /// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147417887)]
		void LostFocus();
	}
}
