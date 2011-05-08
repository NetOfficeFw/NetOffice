using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("AC09","AC10","AC11","AC12","AC14")]
	[ComImport, Guid("F163F201-ADA2-11CF-89A9-00A0C9054129"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _References_Events
	{
		[SupportByLibrary("AC09","AC10","AC11","AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByLibrary("AC09","AC10","AC11","AC12","AC14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _References_Events_SinkHelper : SinkHelper, _References_Events
	{
		#region Static
		
		public static readonly string Id = "F163F201-ADA2-11CF-89A9-00A0C9054129";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _References_Events_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _References_Events Members
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			LateBindingApi.AccessApi.Reference newReference = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reference) as LateBindingApi.AccessApi.Reference;
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

			LateBindingApi.AccessApi.Reference newReference = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, reference) as LateBindingApi.AccessApi.Reference;
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