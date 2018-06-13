using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.NameSpaceEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NameSpaceEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.NameSpaceEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from NameSpaceEvents
        /// </summary>
        public static readonly string Id = "0006308C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public NameSpaceEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region NameSpaceEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pages"></param>
        /// <param name="folder"></param>
        public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("OptionsPagesAdd"))
            {
                Invoker.ReleaseParamsArray(pages, folder);
                return;
            }

			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.PropertyPages>(EventClass, pages, typeof(NetOffice.OutlookApi.PropertyPages));
            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;

            object[] paramsArray = new object[2];
			paramsArray[0] = newPages;
			paramsArray[1] = newFolder;
			EventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void AutoDiscoverComplete()
        {
            if (!Validate("AutoDiscoverComplete"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AutoDiscoverComplete", ref paramsArray);
		}

		#endregion
	}	
}

