using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// NativeInterface IRibbonUI SupportByVersion Office, 12,14,15,16
    /// The object that is returned by the onLoad procedure specified on the customUI tag.
    /// The object contains methods for invalidating control properties and for refreshing the user interface.
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface IRibbonUI
	{
        #region Methods

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		void Invalidate();

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Invalidates the cached value for a single control on the Ribbon user interface.
        /// </summary>
        /// <param name="ControlID">Specifies the ID of the control that will be invalidated</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)]
		void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Used to invalidate a built-in control.
        /// </summary>
        /// <param name="ControlID">Specified the identifier of the control that will be invalidated.</param>
		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)]
		void InvalidateControlMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Activates the specified custom tab.
        /// </summary>
        /// <param name="ControlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(4)]
		void ActivateTab([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Activates the specified built-in tab.
        /// </summary>
        /// <param name="ControlID">Specifies the identifier of the custom Ribbon tab to be activated.</param>
		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(5)]
		void ActivateTabMso([In, MarshalAs(UnmanagedType.BStr)]string ControlID);

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab which includes the identifier and the namespace of the tab.
        /// </summary>
        /// <param name="ControlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
        /// <param name="Namespace">Specifies the namespace of the tab element</param>
		[SupportByVersion("Office", 14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(6)]
		void ActivateTabQ([In, MarshalAs(UnmanagedType.BStr)]string ControlID, [In, MarshalAs(UnmanagedType.BStr)]string Namespace);

		#endregion
	}
}
