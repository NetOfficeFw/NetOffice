using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.EventContracts
{
    /// <summary>
    /// _SubFormEvents
    /// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("BC9E4350-F037-11CD-8701-00AA003F0F07"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _SubFormEvents
	{
        /// <summary>
        /// Enter
        /// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2019)]
		void Enter();

        /// <summary>
        /// Exit
        /// </summary>
        /// <param name="cancel"></param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2075)]
		void Exit([In] [Out] ref object cancel);
	}
}
