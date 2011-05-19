using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("Outlook", 9,10,11,12,14)]
	[ComImport, Guid("00063085-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface SyncObjectEvents
	{
		[SupportByLibrary("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void SyncStart();

		[SupportByLibrary("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void Progress([In] object state, [In] object description, [In] object value, [In] object max);

		[SupportByLibrary("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void OnError([In] object code, [In] object description);

		[SupportByLibrary("Outlook", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void SyncEnd();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class SyncObjectEvents_SinkHelper : SinkHelper, SyncObjectEvents
	{
		#region Static
		
		public static readonly string Id = "00063085-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public SyncObjectEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region SyncObjectEvents Members
		
		public void SyncStart()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SyncStart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Progress([In] object state, [In] object description, [In] object value, [In] object max)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Progress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(state, description, value, max);
				return;
			}

			NetOffice.OutlookApi.Enums.OlSyncState newState = (NetOffice.OutlookApi.Enums.OlSyncState)state;
			string newDescription = (string)description;
			Int32 newValue = (Int32)value;
			Int32 newMax = (Int32)max;
			object[] paramsArray = new object[4];
			paramsArray[0] = newState;
			paramsArray[1] = newDescription;
			paramsArray[2] = newValue;
			paramsArray[3] = newMax;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void OnError([In] object code, [In] object description)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnError");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(code, description);
				return;
			}

			Int32 newCode = (Int32)code;
			string newDescription = (string)description;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCode;
			paramsArray[1] = newDescription;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SyncEnd()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SyncEnd");
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