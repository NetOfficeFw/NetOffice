using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetOffice.Attributes;

// Not in NetOffice.OfficeApi.Native namespace for backward compatibility
namespace NetOffice.OfficeApi
{
    /// <summary>
    /// NativeInterface IRibbonControl SupportByVersion Office, 12,14,15,16
    /// Represents the object passed into every Ribbon user interface (UI) control's callback procedure.
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface IRibbonControl
	{
        #region Properties

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Get
        /// Gets the ID of the control specified in the Ribbon XML markup customization file.
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		[DispId(1)]
		string Id{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)] get;}

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Get
        /// Represents the active window containing the Ribbon user interface that triggers a callback procedure.
        /// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		[DispId(2)]
		object Context{[return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(2)] get;}

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// Get
        /// Used to store arbitrary strings and fetch them at runtime.
        /// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		[DispId(3)]
		string Tag{[return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(3)] get;}

		#endregion
	}
}
