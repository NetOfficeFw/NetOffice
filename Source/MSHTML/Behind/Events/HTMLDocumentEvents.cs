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
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLDocumentEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLDocumentEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLDocumentEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLDocumentEvents
		/// </summary>
		public static readonly string Id = "3050F260-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLDocumentEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region HTMLDocumentEvents
		
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
		public void onclick()
		{
            if (!Validate("onclick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onclick", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondblclick()
		{
            if (!Validate("ondblclick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondblclick", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onkeydown()
		{
            if (!Validate("onkeydown"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onkeydown", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onkeyup()
		{
            if (!Validate("onkeyup"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onkeyup", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onkeypress()
		{
            if (!Validate("onkeypress"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onkeypress", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmousedown()
		{
            if (!Validate("onmousedown"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmousedown", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmousemove()
		{
            if (!Validate("onmousemove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmousemove", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmouseup()
		{
            if (!Validate("onmouseup"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmouseup", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmouseout()
		{
            if (!Validate("onmouseout"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmouseout", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmouseover()
		{
            if (!Validate("onmouseover"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmouseover", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onreadystatechange()
		{
            if (!Validate("onreadystatechange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onreadystatechange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforeupdate()
		{
            if (!Validate("onbeforeupdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onafterupdate()
		{
            if (!Validate("onafterupdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onafterupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowexit()
		{
            if (!Validate("onrowexit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowexit", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowenter()
		{
            if (!Validate("onrowenter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowenter", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondragstart()
		{
            if (!Validate("ondragstart"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondragstart", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onselectstart()
		{
            if (!Validate("onselectstart"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onselectstart", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onerrorupdate()
		{
            if (!Validate("onerrorupdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onerrorupdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void oncontextmenu()
		{
            if (!Validate("oncontextmenu"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("oncontextmenu", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onstop()
		{
            if (!Validate("onstop"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onstop", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowsdelete()
		{
            if (!Validate("onrowsdelete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowsdelete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onrowsinserted()
		{
            if (!Validate("onrowsinserted"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onrowsinserted", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void oncellchange()
		{
            if (!Validate("oncellchange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("oncellchange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onpropertychange()
		{
            if (!Validate("onpropertychange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onpropertychange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondatasetchanged()
		{
            if (!Validate("ondatasetchanged"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondatasetchanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondataavailable()
		{
            if (!Validate("ondataavailable"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondataavailable", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondatasetcomplete()
		{
            if (!Validate("ondatasetcomplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondatasetcomplete", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforeeditfocus()
		{
            if (!Validate("onbeforeeditfocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeeditfocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onselectionchange()
		{
            if (!Validate("onselectionchange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onselectionchange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void oncontrolselect()
		{
            if (!Validate("oncontrolselect"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("oncontrolselect", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onmousewheel()
		{
            if (!Validate("onmousewheel"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onmousewheel", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onfocusin()
		{
            if (!Validate("onfocusin"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onfocusin", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onfocusout()
		{
            if (!Validate("onfocusout"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onfocusout", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onactivate()
		{
            if (!Validate("onactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ondeactivate()
		{
            if (!Validate("ondeactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ondeactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforeactivate()
		{
            if (!Validate("onbeforeactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforeactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void onbeforedeactivate()
		{
            if (!Validate("onbeforedeactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("onbeforedeactivate", ref paramsArray);
		}

		#endregion
	}
	
}
