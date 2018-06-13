using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.OutlookBarGroupsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarGroupsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.OutlookBarGroupsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from OutlookBarGroupsEvents
        /// </summary>
        public static readonly string Id = "0006307B-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OutlookBarGroupsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region OutlookBarGroupsEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="newGroup"></param>
		public void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup)
		{
            if (!Validate("GroupAdd"))
            {
                Invoker.ReleaseParamsArray(newGroup);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newNewGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, newGroup, typeof(NetOffice.OutlookApi.OutlookBarGroup));
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewGroup;
			EventBinding.RaiseCustomEvent("GroupAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cancel"></param>
		public void BeforeGroupAdd([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeGroupAdd"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeGroupAdd", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="group"></param>
        /// <param name="cancel"></param>
		public void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeGroupRemove"))
            {
                Invoker.ReleaseParamsArray(group, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, group, typeof(NetOffice.OutlookApi.OutlookBarGroup));
			object[] paramsArray = new object[2];
			paramsArray[0] = newGroup;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeGroupRemove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		#endregion
	}	
}

