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
	/// Default implementation of <see cref="NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents2"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class WHTMLControlEvents2_SinkHelper : SinkHelper, NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents2
	{
		#region Static
		
		/// <summary>
		/// Interface Id from WHTMLControlEvents2
		/// </summary>
		public static readonly string Id = "47FF8FE1-6198-11CF-8CE8-00AA006CB389";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public WHTMLControlEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region WHTMLControlEvents2
		
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
