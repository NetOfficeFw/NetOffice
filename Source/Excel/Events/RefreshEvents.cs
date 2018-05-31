using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.EventContracts
{
    [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002441B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RefreshEvents
	{
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1596)]
		void BeforeRefresh([In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1597)]
		void AfterRefresh([In] object success);
	}
}