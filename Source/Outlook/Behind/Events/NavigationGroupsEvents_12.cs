using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.NavigationGroupsEvents_12"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NavigationGroupsEvents_12_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.NavigationGroupsEvents_12
	{
        #region Static

        /// <summary>
        /// Interface Id from NavigationGroupsEvents_12
        /// </summary>
        public static readonly string Id = "000630F4-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public NavigationGroupsEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region NavigationGroupsEvents_12
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="navigationFolder"></param>
		public void SelectedChange([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder)
		{
            if (!Validate("SelectedChange"))
            {
                Invoker.ReleaseParamsArray(navigationFolder);
                return;
            }

			NetOffice.OutlookApi.NavigationFolder newNavigationFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationFolder>(EventClass, navigationFolder, typeof(NetOffice.OutlookApi.NavigationFolder));
			object[] paramsArray = new object[1];
			paramsArray[0] = newNavigationFolder;
			EventBinding.RaiseCustomEvent("SelectedChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="navigationFolder"></param>
		public void NavigationFolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder)
		{
            if (!Validate("NavigationFolderAdd"))
            {
                Invoker.ReleaseParamsArray(navigationFolder);
                return;
            }

            NetOffice.OutlookApi.NavigationFolder newNavigationFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationFolder>(EventClass, navigationFolder, typeof(NetOffice.OutlookApi.NavigationFolder));
            object[] paramsArray = new object[1];
			paramsArray[0] = newNavigationFolder;
			EventBinding.RaiseCustomEvent("NavigationFolderAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void NavigationFolderRemove()
        {
            if (!Validate("NavigationFolderRemove"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("NavigationFolderRemove", ref paramsArray);
		}

		#endregion
	}	
}

