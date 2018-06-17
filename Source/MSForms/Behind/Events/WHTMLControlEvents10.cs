using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.MSFormsApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents10"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class WHTMLControlEvents10_SinkHelper : SinkHelper, NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents10
	{
		#region Static
		
		/// <summary>
		/// Interface Id from WHTMLControlEvents10
		/// </summary>
		public static readonly string Id = "47FF8FE9-6198-11CF-8CE8-00AA006CB389";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public WHTMLControlEvents10_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region WHTMLControlEvents10
		
		/// <summary>
		/// 
		/// </summary>
		public void Click()
		{
            if (!Validate("Click"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		#endregion
	}
	
}
