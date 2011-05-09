using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.VBIDEApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("VBE5.3","VBE12")]
	[ComImport, Guid("0002E118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferencesEvents
	{
		[SupportByLibrary("VBE5.3","VBE12")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByLibrary("VBE5.3","VBE12")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferencesEvents_SinkHelper : SinkHelper, _dispReferencesEvents
	{
		#region Static
		
		public static readonly string Id = "0002E118-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _dispReferencesEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _dispReferencesEvents Members
		
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