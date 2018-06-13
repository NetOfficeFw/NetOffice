using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void AccountSelector_SelectedAccountChangeEventHandler(NetOffice.OutlookApi.Account selectedAccount);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass AccountSelector 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867249.aspx </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.AccountSelectorEvents))]
	[TypeId("00061103-0000-0000-C000-000000000046")]
    public interface AccountSelector : _AccountSelector, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869081.aspx </remarks>
        [SupportByVersion("Outlook", 14,15,16)]
		event AccountSelector_SelectedAccountChangeEventHandler SelectedAccountChangeEvent;

        #endregion
    }
}
