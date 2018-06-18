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
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLMapEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLMapEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLMapEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLMapEvents
		/// </summary>
		public static readonly string Id = "3050F3BA-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLMapEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region HTMLMapEvents

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
        public void onfilterchange()
        {
            if (!Validate("onfilterchange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onfilterchange", ref paramsArray);
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
        public void onlosecapture()
        {
            if (!Validate("onlosecapture"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onlosecapture", ref paramsArray);
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
        public void ondrag()
        {
            if (!Validate("ondrag"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondrag", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void ondragend()
        {
            if (!Validate("ondragend"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondragend", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void ondragenter()
        {
            if (!Validate("ondragenter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondragenter", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void ondragover()
        {
            if (!Validate("ondragover"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondragover", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void ondragleave()
        {
            if (!Validate("ondragleave"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondragleave", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void ondrop()
        {
            if (!Validate("ondrop"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ondrop", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onbeforecut()
        {
            if (!Validate("onbeforecut"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onbeforecut", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void oncut()
        {
            if (!Validate("oncut"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("oncut", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onbeforecopy()
        {
            if (!Validate("onbeforecopy"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onbeforecopy", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void oncopy()
        {
            if (!Validate("oncopy"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("oncopy", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onbeforepaste()
        {
            if (!Validate("onbeforepaste"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onbeforepaste", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onpaste()
        {
            if (!Validate("onpaste"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onpaste", ref paramsArray);
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
        public void onlayoutcomplete()
        {
            if (!Validate("onlayoutcomplete"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onlayoutcomplete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onpage()
        {
            if (!Validate("onpage"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onpage", ref paramsArray);
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
        public void onmove()
        {
            if (!Validate("onmove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onmove", ref paramsArray);
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
        public void onmovestart()
        {
            if (!Validate("onmovestart"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onmovestart", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onmoveend()
        {
            if (!Validate("onmoveend"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onmoveend", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onresizestart()
        {
            if (!Validate("onresizestart"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onresizestart", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onresizeend()
        {
            if (!Validate("onresizeend"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onresizeend", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onmouseenter()
        {
            if (!Validate("onmouseenter"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onmouseenter", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
        public void onmouseleave()
        {
            if (!Validate("onmouseleave"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("onmouseleave", ref paramsArray);
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

        #endregion
    }
	
}
