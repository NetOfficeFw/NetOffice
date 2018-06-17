using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// DispEmptyCellEvents
    /// </summary>
	[SupportByVersion("Access", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("3B06E988-E47C-11CD-8701-00AA003F0F07"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DispEmptyCellEvents
	{
	}

}
