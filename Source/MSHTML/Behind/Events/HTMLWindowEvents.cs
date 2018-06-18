using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.MSHTMLApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLWindowEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLWindowEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLWindowEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLWindowEvents
		/// </summary>
		public static readonly string Id = "96A0A4E0-D062-11CF-94B6-00AA0060275C";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLWindowEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLWindowEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void onload()
		{
            if (!Validate("onload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onload", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onunload()
		{
            if (!Validate("onunload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onunload", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onhelp()
		{
            if (!Validate("onhelp"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onhelp", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onfocus()
		{
            if (!Validate("onfocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onfocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onblur()
		{
            if (!Validate("onblur"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onblur", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="description"></param>
		/// <param name="url"></param>
		/// <param name="line"></param>
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

		/// <summary>
		/// 
		/// </summary>
		public void onresize()
		{
            if (!Validate("onresize"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onresize", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onscroll()
		{
            if (!Validate("onscroll"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onscroll", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforeunload()
		{
            if (!Validate("onbeforeunload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeunload", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforeprint()
		{
            if (!Validate("onbeforeprint"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeprint", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
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
	
}
