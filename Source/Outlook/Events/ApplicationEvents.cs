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
    [ComImport, Guid("0006304E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void NewMail();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Startup();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void Quit();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_SinkHelper : SinkHelper, ApplicationEvents
	{
		#region Static
		
		public static readonly string Id = "0006304E-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region ApplicationEvents Members
		
		public void ItemSend([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemSend");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item, cancel);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newItem;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ItemSend", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void NewMail()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewMail");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("NewMail", ref paramsArray);
		}

		public void Reminder([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Reminder");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(item);
				return;
			}

			object newItem = Factory.CreateObjectFromComProxy(_eventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			_eventBinding.RaiseCustomEvent("Reminder", ref paramsArray);
		}

		public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OptionsPagesAdd");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pages);
				return;
			}

			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateObjectFromComProxy(_eventClass, pages) as NetOffice.OutlookApi.PropertyPages;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPages;
			_eventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

		public void Startup()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Startup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Startup", ref paramsArray);
		}

		public void Quit()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Quit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}