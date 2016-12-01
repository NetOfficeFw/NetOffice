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
	[ComImport, Guid("000C0354-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarComboBoxEvents
	{
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarComboBoxEvents_SinkHelper : SinkHelper, _CommandBarComboBoxEvents
	{
		#region Static
		
		public static readonly string Id = "000C0354-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _CommandBarComboBoxEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _CommandBarComboBoxEvents Members
		
		public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Change");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(ctrl);
				return;
			}

			NetOffice.OfficeApi.CommandBarComboBox newCtrl = Factory.CreateObjectFromComProxy(_eventClass, ctrl) as NetOffice.OfficeApi.CommandBarComboBox;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCtrl;
			_eventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}