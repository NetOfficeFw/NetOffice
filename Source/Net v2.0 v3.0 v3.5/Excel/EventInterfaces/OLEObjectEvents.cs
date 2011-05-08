using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
	[ComImport, Guid("00024410-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OLEObjectEvents
	{
		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1541)]
		void GotFocus();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1542)]
		void LostFocus();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OLEObjectEvents_SinkHelper : SinkHelper, OLEObjectEvents
	{
		#region Static
		
		public static readonly string Id = "00024410-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public OLEObjectEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region OLEObjectEvents Members
		
		public void GotFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GotFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void LostFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LostFocus");
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