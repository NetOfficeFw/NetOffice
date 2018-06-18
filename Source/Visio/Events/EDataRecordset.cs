using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.EventContracts
{
    /// <summary>
    /// EDataRecordset
    /// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B11-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EDataRecordset
	{
		/// <summary>
		/// DataRecordsetChanged
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordsetChanged", typeof(VisioApi.IVDataRecordsetChangedEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		/// <summary>
		/// BeforeDataRecordsetDelete
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);
	}

}
