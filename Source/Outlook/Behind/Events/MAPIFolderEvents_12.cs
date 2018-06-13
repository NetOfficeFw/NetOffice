using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.MAPIFolderEvents_12"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MAPIFolderEvents_12_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.MAPIFolderEvents_12
	{
        #region Static

        /// <summary>
        /// Interface Id from MAPIFolderEvents_12
        /// </summary>
        public static readonly string Id = "000630F7-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public MAPIFolderEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region MAPIFolderEvents_12
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="moveTo"></param>
        /// <param name="cancel"></param>
		public void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeFolderMove"))
            {
                Invoker.ReleaseParamsArray(moveTo, cancel);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateEventArgumentObjectFromComProxy(EventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[2];
			paramsArray[0] = newMoveTo;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeFolderMove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="moveTo"></param>
        /// <param name="cancel"></param>
		public void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeItemMove"))
            {
                Invoker.ReleaseParamsArray(item, moveTo, cancel);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
            NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateEventArgumentObjectFromComProxy(EventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[3];
			paramsArray[0] = newItem;
			paramsArray[1] = newMoveTo;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("BeforeItemMove", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

		#endregion
	}
}
