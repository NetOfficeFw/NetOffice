using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OfficeApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
	[ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarButtonEvents
	{
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarButtonEvents_SinkHelper : SinkHelper, _CommandBarButtonEvents
	{
		#region Static
		
		public static readonly string Id = "000C0351-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _CommandBarButtonEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region _CommandBarButtonEvents Members
		
		public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Click");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(ctrl, cancelDefault);
				return;
			}

			NetOffice.OfficeApi.CommandBarButton newCtrl = Factory.CreateObjectFromComProxy(_eventClass, ctrl) as NetOffice.OfficeApi.CommandBarButton;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCtrl;
			paramsArray.SetValue(cancelDefault, 1);
			_eventBinding.RaiseCustomEvent("Click", ref paramsArray);

			cancelDefault = (bool)paramsArray[1];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}