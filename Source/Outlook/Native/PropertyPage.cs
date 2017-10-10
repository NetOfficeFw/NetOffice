using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Native
{
    /// <summary>
    /// NativeInterface PropertyPage SupportByVersion Outlook, 9,10,11,12,14,15,16
    /// Represents a custom property page in the Options dialog box or in the folder Properties dialog box.
    /// </summary>
    [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[ComImport, Guid("0006307E-0000-0000-C000-000000000046"), TypeLibType((short) 4096)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface PropertyPage
    {
        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Returns Help information about a custom property page.
        /// </summary>
        /// <param name="HelpFile">Specifies the Help file associated with the property page</param>
        /// <param name="HelpContext">Specifies the context ID of the Help topic associated with the property page</param>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("HelpFile", SinkArgumentType.String)]
        [SinkArgument("HelpContext", SinkArgumentType.Int32)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8448)]
        void GetPageInfo([MarshalAs(19)] [In] [Out] ref string HelpFile, [In] [Out] ref int HelpContext);
      
        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Get
        /// Returns a Boolean value that indicates whether the contents of a custom property page have been altered.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(8449)]
        bool Dirty { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8449)] get; }

        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Applies the changes that have been made in a custom property page.
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8450)]
		void Apply();
	}
}

