using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.EventContracts
{
    /// <summary>
    /// IRangeEvents
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("B8891063-2B00-48EC-957F-6DEBEADE9D8B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IRangeEvents
	{
        /// <summary>
        /// Change
        /// </summary>
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1510)]
		void Change();
	}
}
