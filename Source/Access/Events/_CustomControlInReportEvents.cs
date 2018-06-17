using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// _CustomControlInReportEvents
    /// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("300471E2-7426-11CE-AB64-00AA0042B7CE"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomControlInReportEvents
	{
	}
}
