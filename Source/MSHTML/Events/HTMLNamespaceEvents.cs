using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi.EventContracts
{
    /// <summary>
    /// HTMLNamespaceEvents
    /// </summary>
	[SupportByVersion("MSHTML", 4)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("3050F6BD-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLNamespaceEvents
	{
		/// <summary>
		/// onreadystatechange
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
        [SinkArgument("pEvtObj", typeof(MSHTMLApi.IHTMLEventObj))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-609)]
		void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj);
	}

}
