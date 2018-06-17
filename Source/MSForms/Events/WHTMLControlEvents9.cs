using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.EventContracts
{
    /// <summary>
    /// WHTMLControlEvents9
    /// </summary>
	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("47FF8FE8-6198-11CF-8CE8-00AA006CB389"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface WHTMLControlEvents9
	{
        /// <summary>
        /// Click
        /// </summary>
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();
	}
}
