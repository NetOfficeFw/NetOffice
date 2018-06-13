using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.InspectorsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class InspectorsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.InspectorsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from InspectorsEvents
        /// </summary>
        public static readonly string Id = "00063079-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public InspectorsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region InspectorsEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="inspector"></param>
		public void NewInspector([In, MarshalAs(UnmanagedType.IDispatch)] object inspector)
        {
            if (!Validate("NewInspector"))
            {
                Invoker.ReleaseParamsArray(inspector);
                return;
            }

            NetOffice.OutlookApi._Inspector newInspector = Factory.CreateEventArgumentObjectFromComProxy(EventClass, inspector) as NetOffice.OutlookApi._Inspector;
            object[] paramsArray = new object[1];
			paramsArray[0] = newInspector;
			EventBinding.RaiseCustomEvent("NewInspector", ref paramsArray);
		}

		#endregion
	}	
}
