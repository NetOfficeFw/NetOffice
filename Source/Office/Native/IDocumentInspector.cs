using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// NativeInterface IDocumentInspector SupportByVersion Office, 12,14,15,16
    /// Inspects a document for specific information items or document properties by using a custom Document Inspector module.
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [ComImport, Guid("000CD706-0000-0000-C000-000000000046"), InterfaceType(1), TypeLibType(256)]
    [EntityType(EntityType.IsNativeInterface)]
    public interface IDocumentInspector
    {
        /// <summary>
        /// Gets information about a custom Document Inspector module.
        /// </summary>
        /// <param name="Name">Represents the name of the module.</param>
        /// <param name="Desc">Represents the description of the module.</param>
        [MethodImpl(4096)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [SinkArgument("Name", SinkArgumentType.String)]
        [SinkArgument("Desc", SinkArgumentType.String)]
        void GetInfo([MarshalAs(UnmanagedType.BStr)] out string Name, [MarshalAs(UnmanagedType.BStr)] out string Desc);

        /// <summary>
        /// Inspects a document for specific information items or document properties by using a custom Document Inspector module.
        /// </summary>
        /// <param name="Doc">An object representing the container document.</param>
        /// <param name="Status">An MsoDocInspectorStatus value that represents the results of the inspection.</param>
        /// <param name="Result">Contains a list of the information items or document properties found in the document.</param>
        /// <param name="Action">Indicates to the user what action to take based on the results of the inspection.</param>
        [MethodImpl(4096)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [SinkArgument("Doc", SinkArgumentType.UnknownProxy)]
        [SinkArgument("Status", SinkArgumentType.Enum, typeof(NetOffice.OfficeApi.Enums.MsoDocInspectorStatus))]
        [SinkArgument("Result", SinkArgumentType.String)]
        [SinkArgument("Action", SinkArgumentType.String)]
        void Inspect([MarshalAs(UnmanagedType.IDispatch)] object Doc, out MsoDocInspectorStatus Status, [MarshalAs(UnmanagedType.BStr)] out string Result, [MarshalAs(UnmanagedType.BStr)] out string Action);

        /// <summary>
        /// Performs some action on specific information items or document properties by using a custom Document Inspector module.
        /// </summary>
        /// <param name="Doc">Specifies nn object representing the container object.</param>
        /// <param name="Hwnd">Specifies the unique identifier of the active document window.</param>
        /// <param name="Status">Specifies an enumeration that indicates the status of the action.</param>
        /// <param name="Result">Contains the results of the action.</param>
        [MethodImpl(4096)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [SinkArgument("Doc", SinkArgumentType.UnknownProxy)]
        [SinkArgument("Hwnd", SinkArgumentType.Int32)]
        [SinkArgument("Status", SinkArgumentType.Enum, typeof(NetOffice.OfficeApi.Enums.MsoDocInspectorStatus))]
        [SinkArgument("Result", SinkArgumentType.String)]
        void Fix([MarshalAs(UnmanagedType.IDispatch)] [In] object Doc, [In] int Hwnd, out MsoDocInspectorStatus Status, [MarshalAs(UnmanagedType.BStr)] out string Result);
    }
}