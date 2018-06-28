using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// NativeInterface IRibbonExtensibility SupportByVersion Office, 12,14,15,16
    /// The interface through which the Ribbon user interface (UI) communicates with a COM add-in to customize the UI.
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
    [EntityType(EntityType.IsNativeInterface)]
	public interface IRibbonExtensibility
	{
        #region Methods

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="RibbonID">The ID for the RibbonX UI</param>
        /// <returns>System.String</returns>
        [SupportByVersion("Office", 12,14,15,16)]
		[return: MarshalAs(UnmanagedType.BStr)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)]string RibbonID);

		#endregion
	}
}