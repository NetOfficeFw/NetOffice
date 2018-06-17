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
	/// Default implementation of <see cref="NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class WHTMLControlEvents_SinkHelper : SinkHelper, NetOffice.MSFormsApi.EventContracts.WHTMLControlEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from WHTMLControlEvents
		/// </summary>
		public static readonly string Id = "796ED650-5FE9-11CF-8D68-00AA00BDCE1D";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public WHTMLControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region WHTMLControlEvents Members
		
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
