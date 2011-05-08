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
	[ComImport, Guid("0006307B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarGroupsEvents
	{
		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupAdd([In] [Out] ref object cancel);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarGroupsEvents_SinkHelper : SinkHelper, OutlookBarGroupsEvents
	{
		#region Static
		
		public static readonly string Id = "0006307B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public OutlookBarGroupsEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region OutlookBarGroupsEvents Members
		
		public void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GroupAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newGroup);
				return;
			}

			LateBindingApi.OutlookApi.OutlookBarGroup newNewGroup = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, newGroup) as LateBindingApi.OutlookApi.OutlookBarGroup;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewGroup;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeGroupAdd([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeGroupAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeGroupRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(group, cancel);
				return;
			}

			LateBindingApi.OutlookApi.OutlookBarGroup newGroup = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, group) as LateBindingApi.OutlookApi.OutlookBarGroup;
			object[] paramsArray = new object[2];
			paramsArray[0] = newGroup;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}