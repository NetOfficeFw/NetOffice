using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.StoresEvents_12"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class StoresEvents_12_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.StoresEvents_12
	{
        #region Static

        /// <summary>
        /// Interface Id from StoresEvents_12
        /// </summary>
        public static readonly string Id = "000630F8-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public StoresEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region StoresEvents_12
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="store"></param>
        /// <param name="cancel"></param>
		public void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeStoreRemove"))
            {
                Invoker.ReleaseParamsArray(store, cancel);
                return;
            }

            NetOffice.OutlookApi._Store newStore = Factory.CreateEventArgumentObjectFromComProxy(EventClass, store) as NetOffice.OutlookApi._Store;
            object[] paramsArray = new object[2];
			paramsArray[0] = newStore;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeStoreRemove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="store"></param>
		public void StoreAdd([In, MarshalAs(UnmanagedType.IDispatch)] object store)
        {
            if (!Validate("BeforeStoreRemove"))
            {
                Invoker.ReleaseParamsArray(store);
                return;
            }

            NetOffice.OutlookApi._Store newStore = Factory.CreateEventArgumentObjectFromComProxy(EventClass, store) as NetOffice.OutlookApi._Store;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStore;
			EventBinding.RaiseCustomEvent("StoreAdd", ref paramsArray);
		}

		#endregion
	}	
}
