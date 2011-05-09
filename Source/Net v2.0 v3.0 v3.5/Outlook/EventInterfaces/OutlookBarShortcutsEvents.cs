using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
	[ComImport, Guid("0006307C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarShortcutsEvents
	{
		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeShortcutAdd([In] [Out] ref object cancel);

		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarShortcutsEvents_SinkHelper : SinkHelper, OutlookBarShortcutsEvents
	{
		#region Static
		
		public static readonly string Id = "0006307C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public OutlookBarShortcutsEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region OutlookBarShortcutsEvents Members
		
		public void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShortcutAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newShortcut);
				return;
			}

			NetOffice.OutlookApi.OutlookBarShortcut newNewShortcut = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, newShortcut) as NetOffice.OutlookApi.OutlookBarShortcut;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewShortcut;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeShortcutAdd([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeShortcutAdd");
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

		public void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeShortcutRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(shortcut, cancel);
				return;
			}

			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, shortcut) as NetOffice.OutlookApi.OutlookBarShortcut;
			object[] paramsArray = new object[2];
			paramsArray[0] = newShortcut;
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