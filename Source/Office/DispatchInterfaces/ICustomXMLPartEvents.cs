using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ICustomXMLPartEvents 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000CDB06-0000-0000-C000-000000000046")]
    public interface ICustomXMLPartEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode newNode</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void NodeAfterInsert(NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="oldParentNode">NetOffice.OfficeApi.CustomXMLNode oldParentNode</param>
        /// <param name="oldNextSibling">NetOffice.OfficeApi.CustomXMLNode oldNextSibling</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void NodeAfterDelete(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode oldParentNode, NetOffice.OfficeApi.CustomXMLNode oldNextSibling, bool inUndoRedo);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode newNode</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void NodeAfterReplace(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo);

        #endregion
    }
}
