using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.EventContracts
{
    /// <summary>
    /// ImageListEvents
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("2C247F22-8591-11D1-B16A-00C0F0283628"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ImageListEvents
	{
	}
}
