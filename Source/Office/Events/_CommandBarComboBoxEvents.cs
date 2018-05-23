using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventInterfaces
{	
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C0354-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarComboBoxEvents
	{
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
        [SinkArgument("ctrl", typeof(OfficeApi.CommandBarComboBox))]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl);
	}
}
