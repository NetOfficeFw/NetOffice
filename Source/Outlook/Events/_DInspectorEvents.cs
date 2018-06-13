using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// _DInspectorEvents
    /// </summary>
	[SupportByVersion("Outlook", 10)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("2D9C6D57-BD3C-4275-BED2-73F0EDC18CCE"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DInspectorEvents
	{
	}
}
