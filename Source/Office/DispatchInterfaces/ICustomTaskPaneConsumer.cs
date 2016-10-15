using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	#pragma warning disable
	///<summary>
	/// DispatchInterface ICustomTaskPaneConsumer SupportByVersionAttribute Office, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[ComImport, ComVisible(true), Guid("000C033E-0000-0000-C000-000000000046"), TypeLibType((short) 4288)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public interface ICustomTaskPaneConsumer
	{
		#region Methods

		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(1)]
		void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] object CTPFactoryInst);

		#endregion

		#region Properties

		#endregion
	}
}
	#pragma warning restore
