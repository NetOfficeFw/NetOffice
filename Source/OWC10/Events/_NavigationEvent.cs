using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.EventContracts
{
    /// <summary>
    /// _NavigationEvent
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("4BD09D02-45CC-11D1-B1D1-006097C97F9B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _NavigationEvent
	{
        /// <summary>
        /// ButtonClick
        /// </summary>
        /// <param name="navButton"></param>
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("navButton", typeof(OWC10Api.Enums.NavButtonEnum))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(740)]
		void ButtonClick([In] object navButton);
	}
}
