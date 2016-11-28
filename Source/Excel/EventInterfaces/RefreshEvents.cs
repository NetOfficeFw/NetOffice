using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[ComImport, Guid("0002441B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RefreshEvents
	{
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1596)]
		void BeforeRefresh([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1597)]
		void AfterRefresh([In] object success);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class RefreshEvents_SinkHelper : SinkHelper, RefreshEvents
	{
		#region Static
		
		public static readonly string Id = "0002441B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public RefreshEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region RefreshEvents Members
		
		public void BeforeRefresh([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeRefresh");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeforeRefresh", ref paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void AfterRefresh([In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterRefresh");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(success);
				return;
			}

			bool newSuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSuccess;
			_eventBinding.RaiseCustomEvent("AfterRefresh", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}