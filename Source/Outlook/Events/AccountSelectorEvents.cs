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

	[SupportByVersion("Outlook", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063104-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface AccountSelectorEvents
	{
		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64627)]
		void SelectedAccountChange([In, MarshalAs(UnmanagedType.IDispatch)] object selectedAccount);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AccountSelectorEvents_SinkHelper : SinkHelper, AccountSelectorEvents
	{
		#region Static
		
		public static readonly string Id = "00063104-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public AccountSelectorEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region AccountSelectorEvents Members
		
		public void SelectedAccountChange([In, MarshalAs(UnmanagedType.IDispatch)] object selectedAccount)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectedAccountChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(selectedAccount);
				return;
			}

			NetOffice.OutlookApi.Account newSelectedAccount = Factory.CreateObjectFromComProxy(_eventClass, selectedAccount) as NetOffice.OutlookApi.Account;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelectedAccount;
			_eventBinding.RaiseCustomEvent("SelectedAccountChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}