using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.AccountSelectorEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AccountSelectorEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.AccountSelectorEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from AccountSelectorEvents
        /// </summary>
        public static readonly string Id = "00063104-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public AccountSelectorEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion
		
		#region AccountSelectorEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="selectedAccount"></param>
		public void SelectedAccountChange([In, MarshalAs(UnmanagedType.IDispatch)] object selectedAccount)
        {
            if (!Validate("SelectedAccountChange"))
            {
                Invoker.ReleaseParamsArray(selectedAccount);
            }

			NetOffice.OutlookApi.Account newSelectedAccount = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Account>(EventClass, selectedAccount, typeof(NetOffice.OutlookApi.Account));
			object[] paramsArray = new object[1];
			paramsArray[0] = newSelectedAccount;
			EventBinding.RaiseCustomEvent("SelectedAccountChange", ref paramsArray);
		}

		#endregion
	}
}

