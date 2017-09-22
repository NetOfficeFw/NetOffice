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
        [SinkArgument("selectedAccount", typeof(OutlookApi.Account))]
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

        #region Ctor

        public AccountSelectorEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region AccountSelectorEvents
		
		public void SelectedAccountChange([In, MarshalAs(UnmanagedType.IDispatch)] object selectedAccount)
        {
            if (!Validate("SelectedAccountChange"))
            {
                Invoker.ReleaseParamsArray(selectedAccount);
            }

			NetOffice.OutlookApi.Account newSelectedAccount = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Account>(EventClass, selectedAccount, NetOffice.OutlookApi.Account.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelectedAccount;
			EventBinding.RaiseCustomEvent("SelectedAccountChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}