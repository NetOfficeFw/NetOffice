using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.ItemsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ItemsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.ItemsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from ItemsEvents
        /// </summary>
        public static readonly string Id = "00063077-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public ItemsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ItemsEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
		public void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
            if (!Validate("ItemAdd"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("ItemAdd", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
		public void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
            if (!Validate("ItemChange"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

            object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
            object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("ItemChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
		public void ItemRemove()
		{
            if (!Validate("ItemRemove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ItemRemove", ref paramsArray);
		}

		#endregion
	}	
}
