using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Outlook", 12,14,15,16)]
	[ComImport, Guid("000630F8-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface StoresEvents_12
	{
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64433)]
		void BeforeStoreRemove([In, MarshalAs(UnmanagedType.IDispatch)] object store, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
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
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
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

			NetOffice.OutlookApi._Store newStore = Factory.CreateObjectFromComProxy(_eventClass, store) as NetOffice.OutlookApi._Store;
			object[] paramsArray = new object[2];
			paramsArray[0] = newStore;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeStoreRemove", ref paramsArray);

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

			NetOffice.OutlookApi._Store newStore = Factory.CreateObjectFromComProxy(_eventClass, store) as NetOffice.OutlookApi._Store;
			object[] paramsArray = new object[1];
			paramsArray[0] = newStore;
			_eventBinding.RaiseCustomEvent("StoreAdd", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}