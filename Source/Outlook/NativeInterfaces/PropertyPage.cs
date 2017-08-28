using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
    /// <summary>
    /// NativeInterface PropertyPage SupportByVersion Outlook, 9,10,11,12,14,15,16
    /// Represents a custom property page in the Options dialog box or in the folder Properties dialog box.
    /// </summary>
    [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("0006307E-0000-0000-C000-000000000046"), TypeLibType((short) 4096)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface PropertyPage
	{
        #region Methods

        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Returns Help information about a custom property page.
        /// </summary>
        /// <param name="HelpFile">Specifies the Help file associated with the property page</param>
        /// <param name="HelpContext">Specifies the context ID of the Help topic associated with the property page</param>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8448)]
		void GetPageInfo([In, MarshalAs(UnmanagedType.BStr)]string HelpFile, [In]Int32 HelpContext);

        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Applies the changes that have been made in a custom property page.
        /// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8450)]
		void Apply();

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Outlook, 9,10,11,12,14,15,16
        /// Get
        /// Returns a Boolean value that indicates whether the contents of a custom property page have been altered.
        /// </summary>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[DispId(8449)]
		bool Dirty{[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8449)] get;}

		#endregion
	}
}

