using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("Outlook", 9,10,11,12,14)]
	[ComImport, Guid("0006308C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NameSpaceEvents
	{
		[SupportByLibraryAttribute("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByLibraryAttribute("Outlook", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64557)]
		void AutoDiscoverComplete();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NameSpaceEvents_SinkHelper : SinkHelper, NameSpaceEvents
	{
		#region Static
		
		public static readonly string Id = "0006308C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public NameSpaceEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region NameSpaceEvents Members
		
		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OptionsPagesAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pages, folder);
				return;
			}

			NetOffice.OutlookApi.PropertyPages newPages = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pages) as NetOffice.OutlookApi.PropertyPages;
			NetOffice.OutlookApi.MAPIFolder newFolder = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, folder) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPages;
			paramsArray[1] = newFolder;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AutoDiscoverComplete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AutoDiscoverComplete");
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