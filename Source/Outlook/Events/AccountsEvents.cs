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
    [ComImport, Guid("00063105-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface AccountsEvents
	{
		[SupportByVersion("Outlook", 14,15,16)]
        [SinkArgument("account", typeof(OutlookApi.Account))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64620)]
		void AutoDiscoverComplete([In, MarshalAs(UnmanagedType.IDispatch)] object account);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AccountsEvents_SinkHelper : SinkHelper, AccountsEvents
	{
		#region Static
		
		public static readonly string Id = "00063105-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public AccountsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region AccountsEvents
		
		public void AutoDiscoverComplete([In, MarshalAs(UnmanagedType.IDispatch)] object account)
		{
            if (!Validate("AutoDiscoverComplete"))
            {
                Invoker.ReleaseParamsArray(account);
                return;
            }

            NetOffice.OutlookApi.Account newAccount = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Account>(EventClass, account, NetOffice.OutlookApi.Account.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newAccount;
			EventBinding.RaiseCustomEvent("AutoDiscoverComplete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}