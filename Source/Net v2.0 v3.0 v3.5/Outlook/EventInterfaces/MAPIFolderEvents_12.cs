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
	[ComImport, Guid("000630F7-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MAPIFolderEvents_12
	{
		[SupportByLibrary("OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64424)]
		void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);

		[SupportByLibrary("OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64425)]
		void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MAPIFolderEvents_12_SinkHelper : SinkHelper, MAPIFolderEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F7-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public MAPIFolderEvents_12_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region MAPIFolderEvents_12 Members
		
		public void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeFolderMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(moveTo, cancel);
				return;
			}

			LateBindingApi.OutlookApi.MAPIFolder newMoveTo = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, moveTo) as LateBindingApi.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[2];
			paramsArray[0] = newMoveTo;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeItemMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item, moveTo, cancel);
				return;
			}

			object newItem = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			LateBindingApi.OutlookApi.MAPIFolder newMoveTo = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, moveTo) as LateBindingApi.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[3];
			paramsArray[0] = newItem;
			paramsArray[1] = newMoveTo;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}