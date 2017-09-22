using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSHTML", 4)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("96A0A4E0-D062-11CF-94B6-00AA0060275C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface HTMLWindowEvents
	{
		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1003)]
		void onload();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1008)]
		void onunload();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418102)]
		void onhelp();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418111)]
		void onfocus();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147418112)]
		void onblur();

		[SupportByVersion("MSHTML", 4)]
        [SinkArgument("description", SinkArgumentType.String)]
        [SinkArgument("url", SinkArgumentType.String)]
        [SinkArgument("line", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1002)]
		void onerror([In] object description, [In] object url, [In] object line);

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1016)]
		void onresize();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1014)]
		void onscroll();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1017)]
		void onbeforeunload();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1024)]
		void onbeforeprint();

		[SupportByVersion("MSHTML", 4)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1025)]
		void onafterprint();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLWindowEvents_SinkHelper : SinkHelper, HTMLWindowEvents
	{
		#region Static
		
		public static readonly string Id = "96A0A4E0-D062-11CF-94B6-00AA0060275C";
		
		#endregion
	
		#region Ctor

		public HTMLWindowEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLWindowEvents
		
		public void onload()
		{
            if (!Validate("onload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onload", ref paramsArray);
		}

		public void onunload()
		{
            if (!Validate("onunload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onunload", ref paramsArray);
		}

		public void onhelp()
		{
            if (!Validate("onhelp"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onhelp", ref paramsArray);
		}

		public void onfocus()
		{
            if (!Validate("onfocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onfocus", ref paramsArray);
		}

		public void onblur()
		{
            if (!Validate("onblur"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onblur", ref paramsArray);
		}

        public void onerror([In] object description, [In] object url, [In] object line)
		{
            if (!Validate("onerror"))
            {
                Invoker.ReleaseParamsArray(description, url, line);
                return;
            }

			string newdescription = ToString(description);
			string newurl = ToString(url);
			Int32 newline = ToInt32(line);
			object[] paramsArray = new object[3];
			paramsArray[0] = newdescription;
			paramsArray[1] = newurl;
			paramsArray[2] = newline;
			EventBinding.RaiseCustomEvent("onerror", ref paramsArray);
		}

		public void onresize()
		{
            if (!Validate("onresize"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onresize", ref paramsArray);
		}

		public void onscroll()
		{
            if (!Validate("onscroll"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onscroll", ref paramsArray);
		}

		public void onbeforeunload()
		{
            if (!Validate("onbeforeunload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeunload", ref paramsArray);
		}

		public void onbeforeprint()
		{
            if (!Validate("onbeforeprint"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeprint", ref paramsArray);
		}

		public void onafterprint()
		{
            if (!Validate("onafterprint"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onafterprint", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}