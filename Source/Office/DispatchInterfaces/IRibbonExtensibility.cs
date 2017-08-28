using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#pragma warning disable
	/// <summary>
	/// DispatchInterface IRibbonExtensibility SupportByVersionAttribute Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IRibbonExtensibility
	{
		#region Methods

		[SupportByVersion("Office", 12,14,15,16)]
		[return: MarshalAs(UnmanagedType.BStr)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)]string RibbonID);

		#endregion

		#region Properties

		#endregion
	}
}
	#pragma warning restore


