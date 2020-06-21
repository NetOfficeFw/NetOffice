using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.DispatchInterfaces.Interop
{
    /// <summary>
    /// CustomXMLPart interface.
    /// </summary>
    /// <remarks>
    /// Special interface for calling AddNode() in early bind fashion.
    /// </remarks>
    [ComImport]
    [Guid("000CDB05-0000-0000-C000-000000000046")]
    [TypeIdentifier]
    [TypeLibType(4304)]
    public interface _CustomXMLPart : _IMsoDispObj
    {
        /// <summary>
        /// Adds a node to the XML tree.
        /// </summary>
        [DispId(1610809352)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void AddNode(
            [MarshalAs(UnmanagedType.Interface), In] CustomXMLNode Parent,
            [MarshalAs(UnmanagedType.BStr), In] string Name = "",
            [MarshalAs(UnmanagedType.BStr), In] string NamespaceURI = "",
            [MarshalAs(UnmanagedType.Interface), In] CustomXMLNode NextSibling = null,
            [In] MsoCustomXMLNodeType NodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement,
            [MarshalAs(UnmanagedType.BStr), In] string NodeValue = "");
    }
}
