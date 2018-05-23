using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CommandBarButton_ClickEventHandler(NetOffice.OfficeApi.CommandBarButton ctrl, ref bool cancelDefault);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CommandBarButton
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865221.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(NetOffice.OfficeApi.EventInterfaces._CommandBarButtonEvents))]
    public interface CommandBarButton : _CommandBarButton, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Office 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864867.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        event CommandBarButton_ClickEventHandler ClickEvent;

        #endregion
    }
}
