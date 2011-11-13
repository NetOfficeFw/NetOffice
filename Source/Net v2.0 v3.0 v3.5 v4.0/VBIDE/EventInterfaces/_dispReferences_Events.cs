using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.VBIDEApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("VBIDE", 5.3,12)]
	[ComImport, Guid("CDDE3804-2064-11CF-867F-00AA005FF34A"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferences_Events
	{
		[SupportByLibraryAttribute("VBIDE", 5.3,12)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByLibraryAttribute("VBIDE", 5.3,12)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferences_Events_SinkHelper : SinkHelper, _dispReferences_Events
	{
		#region Static
		
		public static readonly string Id = "CDDE3804-2064-11CF-867F-00AA005FF34A";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _dispReferences_Events_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _dispReferences_Events Members
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			NetOffice.VBIDEApi.Reference newReference = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.VBIDEApi.Reference;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemRemoved");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			NetOffice.VBIDEApi.Reference newReference = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.VBIDEApi.Reference;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}