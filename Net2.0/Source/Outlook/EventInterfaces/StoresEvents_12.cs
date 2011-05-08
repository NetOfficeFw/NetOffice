using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("OL12","OL14")]
	[ComImport, Guid("000630F8-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface StoresEvents_12
	{
		[SupportByLibrary("OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64433)]
		void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel);

		[SupportByLibrary("OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void StoreAdd([In, MarshalAs(UnmanagedType.IDispatch)] object store);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class StoresEvents_12_SinkHelper : SinkHelper, StoresEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F8-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public StoresEvents_12_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region StoresEvents_12 Members
		
		public void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeStoreRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(store, cancel);
				return;
			}

			LateBindingApi.OutlookApi._Store newStore = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, store) as LateBindingApi.OutlookApi._Store;
			object[] paramsArray = new object[2];
			paramsArray[0] = newStore;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void StoreAdd([In, MarshalAs(UnmanagedType.IDispatch)] object store)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("StoreAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(store);
				return;
			}

			LateBindingApi.OutlookApi._Store newStore = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, store) as LateBindingApi.OutlookApi._Store;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStore;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}