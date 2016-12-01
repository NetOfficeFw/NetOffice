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
	[ComImport, Guid("000630F7-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MAPIFolderEvents_12
	{
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64424)]
		void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64425)]
		void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MAPIFolderEvents_12_SinkHelper : SinkHelper, MAPIFolderEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F7-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public MAPIFolderEvents_12_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region MAPIFolderEvents_12 Members
		
		public void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeFolderMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(moveTo, cancel);
				return;
			}

			NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateObjectFromComProxy(_eventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[2];
			paramsArray[0] = newMoveTo;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeFolderMove", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeItemMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item, moveTo, cancel);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateObjectFromComProxy(_eventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[3];
			paramsArray[0] = newItem;
			paramsArray[1] = newMoveTo;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("BeforeItemMove", ref paramsArray);

			cancel = (bool)paramsArray[2];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}