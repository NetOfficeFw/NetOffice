using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.OutlookBarPaneEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarPaneEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.OutlookBarPaneEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from OutlookBarPaneEvents
        /// </summary>
        public static readonly string Id = "0006307A-0000-0000-C000-000000000046";

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OutlookBarPaneEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region OutlookBarPaneEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="shortcut"></param>
        /// <param name="cancel"></param>
		public void BeforeNavigate([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeNavigate"))
            {
                Invoker.ReleaseParamsArray(shortcut, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, typeof(NetOffice.OutlookApi.OutlookBarShortcut));
			object[] paramsArray = new object[2];
			paramsArray[0] = newShortcut;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeNavigate", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toGroup"></param>
        /// <param name="cancel"></param>
		public void BeforeGroupSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object toGroup, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeGroupSwitch"))
            {
                Invoker.ReleaseParamsArray(toGroup, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newToGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, toGroup, typeof(NetOffice.OutlookApi.OutlookBarGroup));
			object[] paramsArray = new object[2];
			paramsArray[0] = newToGroup;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeGroupSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

		#endregion
	}	
}

