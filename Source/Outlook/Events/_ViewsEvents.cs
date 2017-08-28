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

	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630A5-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _ViewsEvents
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64071)]
		void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ViewsEvents_SinkHelper : SinkHelper, _ViewsEvents
	{
		#region Static
		
		public static readonly string Id = "000630A5-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public _ViewsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _ViewsEvents Members
		
		public void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(view);
				return;
			}

			NetOffice.OutlookApi.View newView = Factory.CreateObjectFromComProxy(_eventClass, view) as NetOffice.OutlookApi.View;
			object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			_eventBinding.RaiseCustomEvent("ViewAdd", ref paramsArray);
		}

		public void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewRemove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(view);
				return;
			}

			NetOffice.OutlookApi.View newView = Factory.CreateObjectFromComProxy(_eventClass, view) as NetOffice.OutlookApi.View;
			object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			_eventBinding.RaiseCustomEvent("ViewRemove", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}