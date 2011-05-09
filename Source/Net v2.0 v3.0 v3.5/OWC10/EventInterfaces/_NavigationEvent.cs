using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.OWC10Api
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("XWC1.0")]
	[ComImport, Guid("4BD09D02-45CC-11D1-B1D1-006097C97F9B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _NavigationEvent
	{
		[SupportByLibrary("XWC1.0")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(740)]
		void ButtonClick([In] object navButton);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _NavigationEvent_SinkHelper : SinkHelper, _NavigationEvent
	{
		#region Static
		
		public static readonly string Id = "4BD09D02-45CC-11D1-B1D1-006097C97F9B";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _NavigationEvent_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _NavigationEvent Members
		
		public void ButtonClick([In] object navButton)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ButtonClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(navButton);
				return;
			}

			NetOffice.OWC10Api.Enums.NavButtonEnum newNavButton = (NetOffice.OWC10Api.Enums.NavButtonEnum)navButton;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNavButton;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}