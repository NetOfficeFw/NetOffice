using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ExplorersEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorersEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ExplorersEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from ExplorersEvents
        /// </summary>
        public static readonly string Id = "00063078-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ExplorersEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ExplorersEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="explorer"></param>
		public void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer)
		{
            if (!Validate("NewExplorer"))
            {
                Invoker.ReleaseParamsArray(explorer);
                return;
            }

			NetOffice.OutlookApi._Explorer newExplorer = Factory.CreateEventArgumentObjectFromComProxy(EventClass, explorer) as NetOffice.OutlookApi._Explorer;
			object[] paramsArray = new object[1];
			paramsArray[0] = newExplorer;
			EventBinding.RaiseCustomEvent("NewExplorer", ref paramsArray);
		}

		#endregion
	}	
}
