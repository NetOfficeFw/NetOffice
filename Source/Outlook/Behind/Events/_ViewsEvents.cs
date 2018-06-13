using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts._ViewsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ViewsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts._ViewsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from _ViewsEvents
        /// </summary>
        public static readonly string Id = "000630A5-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public _ViewsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}

        #endregion

        #region _ViewsEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="view"></param>
        public void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewAdd"))
            {
                Invoker.ReleaseParamsArray(view);
            }

			NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, typeof(NetOffice.OutlookApi.View));
			object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			EventBinding.RaiseCustomEvent("ViewAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="view"></param>
		public void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewRemove"))
            {
                Invoker.ReleaseParamsArray(view);
            }

            NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, typeof(NetOffice.OutlookApi.View));
            object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			EventBinding.RaiseCustomEvent("ViewRemove", ref paramsArray);
		}

		#endregion
	}
}

