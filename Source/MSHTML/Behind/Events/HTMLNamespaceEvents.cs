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
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLNamespaceEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLNamespaceEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLNamespaceEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLNamespaceEvents
		/// </summary>
		public static readonly string Id = "3050F6BD-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLNamespaceEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLNamespaceEvents
 
		/// <summary>
		/// 
		/// </summary>
		/// <param name="pEvtObj"></param>
		public void onreadystatechange([In, MarshalAs(UnmanagedType.IDispatch)] object pEvtObj)
		{
            if (!Validate("onreadystatechange"))
            {
                return;
            }

			NetOffice.MSHTMLApi.IHTMLEventObj newpEvtObj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLEventObj>(EventClass, pEvtObj, typeof(NetOffice.MSHTMLApi.IHTMLEventObj));
			object[] paramsArray = new object[1];
			paramsArray[0] = newpEvtObj;
			EventBinding.RaiseCustomEvent("onreadystatechange", ref paramsArray);
		}

		#endregion
	}
	
}

