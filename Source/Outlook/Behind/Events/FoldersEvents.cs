using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.FoldersEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class FoldersEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.FoldersEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from FoldersEvents
        /// </summary>
        public static readonly string Id = "00063076-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public FoldersEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region FoldersEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="folder"></param>
		public void FolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object folder)
        {
            if (!Validate("FolderAdd"))
            {
                Invoker.ReleaseParamsArray(folder);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newFolder;
			EventBinding.RaiseCustomEvent("FolderAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="folder"></param>
		public void FolderChange([In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("FolderChange"))
            {
                Invoker.ReleaseParamsArray(folder);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newFolder;
			EventBinding.RaiseCustomEvent("FolderChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void FolderRemove()
		{
            if (!Validate("FolderRemove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("FolderRemove", ref paramsArray);
		}

		#endregion
	}	
}
