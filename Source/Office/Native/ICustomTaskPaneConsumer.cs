using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// NativeInterface ICustomTaskPaneConsumer SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C033E-0000-0000-C000-000000000046"), TypeLibType((short) 4288)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface ICustomTaskPaneConsumer
	{
        #region Methods

        /// <summary>
        /// Passes an ICTPFactory object to a Microsoft ActiveX add-in that can then be used when creating a custom task pane.
        /// </summary>
        /// <param name="CTPFactoryInst">The object is used by an add-in to create a task pane.</param>
        [SupportByVersion("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] object CTPFactoryInst);

		#endregion
	}
}
