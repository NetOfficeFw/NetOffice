using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.EventContracts
{
    /// <summary>
    /// 
    /// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002E131-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispCommandBarControlEvents
	{
        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBarControl"></param>
        /// <param name="handled"></param>
        /// <param name="cancelDefault"></param>
        [SinkArgument("commandBarControl", SinkArgumentType.UnknownProxy)]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [SupportByVersion("VBIDE", 12,14,5.3)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault);
	}
}
