using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006307B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarGroupsEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupAdd([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarGroupsEvents_SinkHelper : SinkHelper, OutlookBarGroupsEvents
	{
		#region Static
		
		public static readonly string Id = "0006307B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public OutlookBarGroupsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region OutlookBarGroupsEvents Members
		
		public void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GroupAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(newGroup);
				return;
			}

			NetOffice.OutlookApi.OutlookBarGroup newNewGroup = Factory.CreateObjectFromComProxy(_eventClass, newGroup) as NetOffice.OutlookApi.OutlookBarGroup;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewGroup;
			_eventBinding.RaiseCustomEvent("GroupAdd", ref paramsArray);
		}

		public void BeforeGroupAdd([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeGroupAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeforeGroupAdd", ref paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeGroupRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(group, cancel);
				return;
			}

			NetOffice.OutlookApi.OutlookBarGroup newGroup = Factory.CreateObjectFromComProxy(_eventClass, group) as NetOffice.OutlookApi.OutlookBarGroup;
			object[] paramsArray = new object[2];
			paramsArray[0] = newGroup;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeGroupRemove", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}