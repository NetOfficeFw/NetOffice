using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.AccessApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._DummyEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DummyEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._DummyEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _DummyEvents
		/// </summary>
		public static readonly string Id = "58BF3100-B580-11CF-89A8-00A0C9054129";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _DummyEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _DummyEvents
		
		/// <summary>
		/// 
		/// </summary>
		public void Initialize()
        {
            if (!Validate("Initialize"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Initialize", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Terminate()
        {
            if (!Validate("Terminate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Terminate", ref paramsArray);
		}

		#endregion
	}
	
}
