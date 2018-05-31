using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CommandBarComboBox_ChangeEventHandler(NetOffice.OfficeApi.CommandBarComboBox ctrl);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CommandBarComboBox
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865547.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(NetOffice.OfficeApi.EventContracts._CommandBarComboBoxEvents))]
    public interface CommandBarComboBox : _CommandBarComboBox, IEventBinding
    {
        #region Events

        /// <summary>
		/// SupportByVersion Office 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864955.aspx </remarks>
		[SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        event CommandBarComboBox_ChangeEventHandler ChangeEvent;

        #endregion
    }
}
