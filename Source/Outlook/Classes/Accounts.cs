using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
    #region Delegates

    #pragma warning disable
    public delegate void Accounts_AutoDiscoverCompleteEventHandler(NetOffice.OutlookApi.Account account);
    #pragma warning restore

    #endregion
    
    /// <summary>
    /// CoClass Accounts 
    /// SupportByVersion Outlook, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862476.aspx </remarks>
    [SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.AccountsEvents))]
	[TypeId("000610C4-0000-0000-C000-000000000046")]
    public interface Accounts : _Accounts, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867339.aspx </remarks>
        [SupportByVersion("Outlook", 14,15,16)]
		event Accounts_AutoDiscoverCompleteEventHandler AutoDiscoverCompleteEvent;

        #endregion
    }
}
