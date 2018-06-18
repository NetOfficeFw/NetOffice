using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi.EventContracts
{
    /// <summary>
    /// HTMLObjectElementEvents
    /// </summary>
	[SupportByVersion("MSHTML", 4)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("3050F3C4-98B5-11CF-BB82-00AA00BDCE0B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLObjectElementEvents
	{
		/// <summary>
		/// onbeforeupdate
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418108)]
		void onbeforeupdate();

		/// <summary>
		/// onafterupdate
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418107)]
		void onafterupdate();

		/// <summary>
		/// onerrorupdate
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418099)]
		void onerrorupdate();

		/// <summary>
		/// onrowexit
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418106)]
		void onrowexit();

		/// <summary>
		/// onrowenter
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418105)]
		void onrowenter();

		/// <summary>
		/// ondatasetchanged
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418098)]
		void ondatasetchanged();

		/// <summary>
		/// ondataavailable
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418097)]
		void ondataavailable();

		/// <summary>
		/// ondatasetcomplete
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418096)]
		void ondatasetcomplete();

		/// <summary>
		/// onerror
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418093)]
		void onerror();

		/// <summary>
		/// onrowsdelete
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418080)]
		void onrowsdelete();

		/// <summary>
		/// onrowsinserted
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418079)]
		void onrowsinserted();

		/// <summary>
		/// oncellchange
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418078)]
		void oncellchange();

		/// <summary>
		/// onreadystatechange
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418092)]
		void onreadystatechange();
	}

}
