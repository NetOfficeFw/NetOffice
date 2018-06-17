using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// _DummyEvents
    /// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("58BF3100-B580-11CF-89A8-00A0C9054129"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DummyEvents
	{
        /// <summary>
        /// Initialize
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-6)]
		void Initialize();

        /// <summary>
        /// Terminate
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-7)]
		void Terminate();
	}
}
