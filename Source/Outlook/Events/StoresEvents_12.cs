using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F8-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface StoresEvents_12
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("store", typeof(NetOffice.OutlookApi._Store))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64433)]
		void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("store", typeof(NetOffice.OutlookApi._Store))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void StoreAdd([In, MarshalAs(UnmanagedType.IDispatch)] object store);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class StoresEvents_12_SinkHelper : SinkHelper, StoresEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F8-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public StoresEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region StoresEvents_12
		
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
	
	#endregion
	
	#pragma warning restore
}