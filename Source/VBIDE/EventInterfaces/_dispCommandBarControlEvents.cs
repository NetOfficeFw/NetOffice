using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VBIDEApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[ComImport, Guid("0002E131-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispCommandBarControlEvents
	{
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispCommandBarControlEvents_SinkHelper : SinkHelper, _dispCommandBarControlEvents
	{
		#region Static
		
		public static readonly string Id = "0002E131-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _dispCommandBarControlEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _dispCommandBarControlEvents Members
		
		public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Click");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(commandBarControl, handled, cancelDefault);
				return;
			}

			object newCommandBarControl = Factory.CreateObjectFromComProxy(_eventClass, commandBarControl) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newCommandBarControl;
			paramsArray.SetValue(handled, 1);
			paramsArray.SetValue(cancelDefault, 2);
			_eventBinding.RaiseCustomEvent("Click", ref paramsArray);

			handled = (bool)paramsArray[1];
			cancelDefault = (bool)paramsArray[2];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}