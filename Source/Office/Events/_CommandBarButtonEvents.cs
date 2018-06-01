using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventContracts
{
    /// <summary>
    /// _CommandBarButtonEvents
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface _CommandBarButtonEvents
    {
        /// <summary>
        /// Click
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [SinkArgument("ctrl", typeof(CommandBarButton))]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault);
    }
}
