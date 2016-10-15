using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// Interface PropertyPage SupportByVersionAttribute Outlook, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("0006307E-0000-0000-C000-000000000046"), TypeLibType((short) 4096)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public interface PropertyPage
	{
		#region Methods

        /// <summary>
        /// 
        /// </summary>
        /// <param name="HelpFile"></param>
        /// <param name="HelpContext"></param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8448)]
		void GetPageInfo([In, MarshalAs(UnmanagedType.BStr)]string HelpFile, [In]Int32 HelpContext);

        /// <summary>
        /// 
        /// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8450)]
		void Apply();

		#endregion

		#region Properties

        /// <summary>
        /// 
        /// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		[DispId(8449)]
		bool Dirty{[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(8449)] get;}

		#endregion
	}
}