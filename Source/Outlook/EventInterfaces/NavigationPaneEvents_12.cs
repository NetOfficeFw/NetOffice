using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.OutlookApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Outlook", 12,14,15,16)]
	[ComImport, Guid("000630F3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NavigationPaneEvents_12
	{
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64457)]
		void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NavigationPaneEvents_12_SinkHelper : SinkHelper, NavigationPaneEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F3-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public NavigationPaneEvents_12_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region NavigationPaneEvents_12 Members
		
		public void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ModuleSwitch");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(currentModule);
				return;
			}

			NetOffice.OutlookApi.NavigationModule newCurrentModule = Factory.CreateObjectFromComProxy(_eventClass, currentModule) as NetOffice.OutlookApi.NavigationModule;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCurrentModule;
			_eventBinding.RaiseCustomEvent("ModuleSwitch", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}