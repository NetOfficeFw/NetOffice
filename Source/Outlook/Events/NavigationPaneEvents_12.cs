using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.EventContracts
{
    /// <summary>
    /// NavigationPaneEvents_12
    /// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NavigationPaneEvents_12
	{
        /// <summary>
        /// ModuleSwitch
        /// </summary>
        /// <param name="currentModule"></param>
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("currentModule", typeof(OutlookApi.NavigationModule))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64457)]
		void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule);
	}
}
