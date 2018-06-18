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
	/// Default implementation of <see cref="NetOffice.MSHTMLApi.EventContracts.HTMLXMLHttpRequestEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class HTMLXMLHttpRequestEvents_SinkHelper : SinkHelper, NetOffice.MSHTMLApi.EventContracts.HTMLXMLHttpRequestEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from HTMLXMLHttpRequestEvents
		/// </summary>
		public static readonly string Id = "30510498-98B5-11CF-BB82-00AA00BDCE0B";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public HTMLXMLHttpRequestEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region HTMLXMLHttpRequestEvents Members
		
		/// <summary>
		/// 
		/// </summary>
		public void ontimeout()
		{
            if (!Validate("ontimeout"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ontimeout", ref paramsArray);
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

		#endregion
	}
	
}
