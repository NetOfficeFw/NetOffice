using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.AccountsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AccountsEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.AccountsEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from AccountsEvents
        /// </summary>
        public static readonly string Id = "00063105-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public AccountsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion
		
		#region AccountsEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="account"></param>
		public void AutoDiscoverComplete([In, MarshalAs(UnmanagedType.IDispatch)] object account)
		{
            if (!Validate("AutoDiscoverComplete"))
            {
                Invoker.ReleaseParamsArray(account);
                return;
            }

            NetOffice.OutlookApi.Account newAccount = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Account>(EventClass, account, typeof(NetOffice.OutlookApi.Account));
			object[] paramsArray = new object[1];
			paramsArray[0] = newAccount;
			EventBinding.RaiseCustomEvent("AutoDiscoverComplete", ref paramsArray);
		}

		#endregion
	}
}

