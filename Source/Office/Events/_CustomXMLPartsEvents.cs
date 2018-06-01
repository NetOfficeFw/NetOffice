using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventContracts
{
    /// <summary>
    /// _CustomXMLPartsEvents
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000CDB0B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartsEvents
	{
        /// <summary>
        /// PartAfterAdd
        /// </summary>
        /// <param name="newPart"></param>
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("newPart", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart);

        /// <summary>
        /// PartBeforeDelete
        /// </summary>
        /// <param name="oldPart"></param>
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("oldPart", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart);

        /// <summary>
        /// PartAfterLoad
        /// </summary>
        /// <param name="part"></param>
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("part", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part);
	}
}
