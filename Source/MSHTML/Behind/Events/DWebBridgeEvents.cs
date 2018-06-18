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
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.DWebBridgeEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DWebBridgeEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.DWebBridgeEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DWebBridgeEvents
		/// </summary>
		public static readonly string Id = "A6D897FF-0A95-11D1-B0BA-006008166E11";
		
		#endregion

		#region Construction

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DWebBridgeEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region DWebBridgeEvents Members
		       
		/// <summary>
		/// 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="eventData"></param>
		public void onscriptletevent([In] object name, [In] object eventData)
		{
            if (!Validate("onscriptletevent"))
            {
                Invoker.ReleaseParamsArray(name, eventData);
                return;
            }

			string newname = ToString(name);
			object neweventData = (object)eventData;
			object[] paramsArray = new object[2];
			paramsArray[0] = newname;
			paramsArray[1] = neweventData;
			EventBinding.RaiseCustomEvent("onscriptletevent", ref paramsArray);
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

		#endregion
	}
	
}
