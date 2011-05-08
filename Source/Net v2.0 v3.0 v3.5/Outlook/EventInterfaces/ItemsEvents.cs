using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
	[ComImport, Guid("00063077-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ItemsEvents
	{
		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void ItemRemove();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ItemsEvents_SinkHelper : SinkHelper, ItemsEvents
	{
		#region Static
		
		public static readonly string Id = "00063077-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ItemsEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ItemsEvents Members
		
		public void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item);
				return;
			}

			object newItem = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item);
				return;
			}

			object newItem = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ItemRemove()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}