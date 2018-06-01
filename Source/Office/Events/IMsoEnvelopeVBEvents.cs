using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventContracts
{
    /// <summary>
    /// IMsoEnvelopeVBEvents
    /// </summary>
    [SupportByVersion("Office", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000672AD-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IMsoEnvelopeVBEvents
	{
        /// <summary>
        /// EnvelopeShow
        /// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void EnvelopeShow();

        /// <summary>
        /// EnvelopeHide
        /// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void EnvelopeHide();
	}
}
