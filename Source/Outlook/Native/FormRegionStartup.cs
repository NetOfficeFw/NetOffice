using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Native
{
    /// <summary>
    /// NativeInterface FormRegionStartup SupportByVersion Outlook, 12,14,15,16 
    /// Defines an interface that allows an add-in to specify the storage and the user interface of a form region, 
    /// obtains an object for that form region, and determines when the form region is about to be displayed in a form or in the Reading Pane.
    /// </summary>
    [SupportByVersion("Outlook", 12, 14, 15, 16)]
    [ComImport, Guid("00063059-0000-0000-C000-000000000046"), TypeLibType(4160)]
    [EntityType(EntityType.IsNativeInterface)]
    public interface FormRegionStartup
    {
        /// <summary>
        /// Obtains appropriate storage for a form region based on the specified information.
        /// </summary>
        /// <param name="FormRegionName">The internal name of the form region. This can be indicated by the name tag in the corresponding form region XML manifest.</param>
        /// <param name="Item">The Outlook item object that caused the loading of the form region.</param>
        /// <param name="LCID">The current locale ID.</param>
        /// <param name="FormRegionMode">The mode that the form region is being loaded into.</param>
        /// <param name="FormRegionSize">The type of form region being loaded, either adjoining or separate.</param>
        /// <returns></returns>
        [DispId(64310)]
        [MethodImpl(4096)]
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [SinkArgument("FormRegionName", SinkArgumentType.String)]
        [SinkArgument("Item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("LCID", SinkArgumentType.Int32)]
        [SinkArgument("OlFormRegionMode", SinkArgumentType.Enum, typeof(NetOffice.OutlookApi.Enums.OlFormRegionMode))]
        [SinkArgument("FormRegionSize", SinkArgumentType.Enum, typeof(NetOffice.OutlookApi.Enums.OlFormRegionSize))]
        [return: MarshalAs(27)]
        object GetFormRegionStorage([MarshalAs(UnmanagedType.BStr)] [In] object FormRegionName, [MarshalAs(UnmanagedType.IDispatch)] [In] object Item, [In] object LCID, [In] object FormRegionMode, [In] object FormRegionSize);

        /// <summary>
        /// Allows an add-in to update the user interface of a form region before it is displayed. 
        /// </summary>
        /// <param name="FormRegion">The FormRegion object representing the form region that is to be displayed</param>
        [DispId(64317)]
        [MethodImpl(4096)]
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [SinkArgument("FormRegion", typeof(FormRegion))]
        void BeforeFormRegionShow([MarshalAs(28)] [In] object FormRegion);

        /// <summary>
        /// Obtains the XML manifest for a form region.
        /// </summary>
        /// <param name="FormRegionName">The name of the form region which is the name used when registering the form region in the Windows registry. </param>
        /// <param name="LCID">The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.</param>
        /// <returns></returns>
        [DispId(64563)]
        [MethodImpl(4096)]
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [SinkArgument("FormRegionName", SinkArgumentType.String)]
        [SinkArgument("LCID", SinkArgumentType.Int32)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object GetFormRegionManifest([MarshalAs(UnmanagedType.BStr)] [In] string FormRegionName, [In] int LCID);
       
        /// <summary>
        /// Obtains an icon image that will be displayed for a particular type of icon for the form region.
        /// </summary>
        /// <param name="FormRegionName">The name of the form region which is the name used when registering the form region in the Windows registry. </param>
        /// <param name="LCID">The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.</param>
        /// <param name="Icon">A constant that identifies the type of icon.</param>
        /// <returns></returns>
        [DispId(64564)]
        [MethodImpl(4096)]
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [SinkArgument("FormRegionName", SinkArgumentType.String)]
        [SinkArgument("LCID", SinkArgumentType.Int32)]
        [SinkArgument("Icon", SinkArgumentType.Enum, typeof(NetOffice.OutlookApi.Enums.OlFormRegionIcon))]
        [return: MarshalAs(27)]
        object GetFormRegionIcon([MarshalAs(UnmanagedType.BStr)] [In] object FormRegionName, [In] object LCID, [In] object Icon);
    }
}
