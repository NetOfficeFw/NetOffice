using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void CustomXMLPart_NodeAfterInsertEventHandler(NetOffice.OfficeApi.CustomXMLNode newNode, bool InUndoRedo);
	public delegate void CustomXMLPart_NodeAfterDeleteEventHandler(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode oldParentNode, NetOffice.OfficeApi.CustomXMLNode oldNextSibling, bool inUndoRedo);
	public delegate void CustomXMLPart_NodeAfterReplaceEventHandler(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass CustomXMLPart
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863497.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.OfficeApi.EventContracts._CustomXMLPartEvents))]
	[TypeId("000CDB08-0000-0000-C000-000000000046")]
    public interface CustomXMLPart : _CustomXMLPart, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Office 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861395.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLPart_NodeAfterInsertEventHandler NodeAfterInsertEvent;

        /// <summary>
		/// SupportByVersion Office 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861395.aspx </remarks>
		[SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLPart_NodeAfterDeleteEventHandler NodeAfterDeleteEvent;

        /// <summary>
		/// SupportByVersion Office 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863732.aspx </remarks>
		[SupportByVersion("Office", 12, 14, 15, 16)]
        event CustomXMLPart_NodeAfterReplaceEventHandler NodeAfterReplaceEvent;

        #endregion
    }
}
