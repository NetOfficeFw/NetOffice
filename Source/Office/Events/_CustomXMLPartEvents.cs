using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.EventContracts
{	
	[SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000CDB07-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartEvents
	{
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("newNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]       
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NodeAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);

		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("oldNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("oldParentNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("oldNextSibling", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void NodeAfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldParentNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldNextSibling, [In] object inUndoRedo);

		[SupportByVersion("Office", 12,14,15,16)]

        [SinkArgument("oldNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("newNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void NodeAfterReplace([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);
	}
}
