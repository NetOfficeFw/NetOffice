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
	[ComImport, Guid("00063079-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface InspectorsEvents
	{
		[SupportByLibrary("OL09","OL10","OL11","OL12","OL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void NewInspector([In, MarshalAs(UnmanagedType.IDispatch)] object inspector);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class InspectorsEvents_SinkHelper : SinkHelper, InspectorsEvents
	{
		#region Static
		
		public static readonly string Id = "00063079-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public InspectorsEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region InspectorsEvents Members
		
		public void NewInspector([In, MarshalAs(UnmanagedType.IDispatch)] object inspector)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewInspector");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(inspector);
				return;
			}

			LateBindingApi.OutlookApi._Inspector newInspector = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, inspector) as LateBindingApi.OutlookApi._Inspector;
			object[] paramsArray = new object[1];
			paramsArray[0] = newInspector;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}