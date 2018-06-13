using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// _DPageWrapCtrlEvents
    /// </summary>
	[SupportByVersion("Outlook", 10)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("494F0971-DD96-11D2-AF70-006008AFF117"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DPageWrapCtrlEvents
	{
	}
}
