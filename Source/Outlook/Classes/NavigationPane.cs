using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void NavigationPane_ModuleSwitchEventHandler(NetOffice.OutlookApi.NavigationModule currentModule);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass NavigationPane 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868696.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.NavigationPaneEvents_12))]
	[TypeId("000610F3-0000-0000-C000-000000000046")]
    public interface NavigationPane : _NavigationPane, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865854.aspx </remarks>
        [SupportByVersion("Outlook", 12,14,15,16)]
		event NavigationPane_ModuleSwitchEventHandler ModuleSwitchEvent;

        #endregion
    }
}
