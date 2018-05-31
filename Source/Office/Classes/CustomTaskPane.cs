using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CustomTaskPane_VisibleStateChangeEventHandler(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst);
	public delegate void CustomTaskPane_DockPositionStateChangeEventHandler(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CustomTaskPane
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862782.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(NetOffice.OfficeApi.EventContracts._CustomTaskPaneEvents))]
    public interface CustomTaskPane : _CustomTaskPane, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Office 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862422.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomTaskPane_VisibleStateChangeEventHandler VisibleStateChangeEvent;

        /// <summary>
        /// SupportByVersion Office 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865561.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomTaskPane_DockPositionStateChangeEventHandler DockPositionStateChangeEvent;

        #endregion
    }
}
