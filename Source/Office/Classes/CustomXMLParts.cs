using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CustomXMLParts_PartAfterAddEventHandler(NetOffice.OfficeApi.CustomXMLPart newPart);
	public delegate void CustomXMLParts_PartBeforeDeleteEventHandler(NetOffice.OfficeApi.CustomXMLPart oldPart);
	public delegate void CustomXMLParts_PartAfterLoadEventHandler(NetOffice.OfficeApi.CustomXMLPart part);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CustomXMLParts
    /// SupportByVersion Office 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863162.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.OfficeApi.EventContracts._CustomXMLPartsEvents))]
	[TypeId("000CDB0C-0000-0000-C000-000000000046")]
    public interface CustomXMLParts : _CustomXMLParts, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864147.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLParts_PartAfterAddEventHandler PartAfterAddEvent;

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861735.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLParts_PartBeforeDeleteEventHandler PartBeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864879.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLParts_PartAfterLoadEventHandler PartAfterLoadEvent;

        #endregion
    }
}
