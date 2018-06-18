using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi.EventContracts
{
    /// <summary>
    /// HTMLWindowEvents
    /// </summary>
	[SupportByVersion("MSHTML", 4)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("96A0A4E0-D062-11CF-94B6-00AA0060275C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLWindowEvents
	{
		/// <summary>
		/// onload
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void onload();

		/// <summary>
		/// onunload
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void onunload();

		/// <summary>
		/// onhelp
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418102)]
		void onhelp();

		/// <summary>
		/// onfocus
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418111)]
		void onfocus();

		/// <summary>
		/// onblur
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418112)]
		void onblur();

		/// <summary>
		/// onerror
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("line", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void onerror([In] object description, [In] object url, [In] object line);

		/// <summary>
		/// onresize
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1016)]
		void onresize();

		/// <summary>
		/// onscroll
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1014)]
		void onscroll();

		/// <summary>
		/// onbeforeunload
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1017)]
		void onbeforeunload();

		/// <summary>
		/// onbeforeprint
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1024)]
		void onbeforeprint();

		/// <summary>
		/// onafterprint
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1025)]
		void onafterprint();
	}

}
