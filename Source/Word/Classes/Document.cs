using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Document_NewEventHandler();
	public delegate void Document_OpenEventHandler();
	public delegate void Document_CloseEventHandler();
	public delegate void Document_SyncEventHandler(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Document_XMLAfterInsertEventHandler(NetOffice.WordApi.XMLNode nNewXMLNode, bool inUndoRedo);
	public delegate void Document_XMLBeforeDeleteEventHandler(NetOffice.WordApi.Range deletedRange, NetOffice.WordApi.XMLNode oldXMLNode, bool inUndoRedo);
	public delegate void Document_ContentControlAfterAddEventHandler(NetOffice.WordApi.ContentControl newContentControl, bool inUndoRedo);
	public delegate void Document_ContentControlBeforeDeleteEventHandler(NetOffice.WordApi.ContentControl oldContentControl, bool inUndoRedo);
	public delegate void Document_ContentControlOnExitEventHandler(NetOffice.WordApi.ContentControl contentControl, ref bool cancel);
	public delegate void Document_ContentControlOnEnterEventHandler(NetOffice.WordApi.ContentControl contentControl);
	public delegate void Document_ContentControlBeforeStoreUpdateEventHandler(NetOffice.WordApi.ContentControl contentControl, ref string content);
	public delegate void Document_ContentControlBeforeContentUpdateEventHandler(NetOffice.WordApi.ContentControl contentControl, ref string content);
	public delegate void Document_BuildingBlockInsertEventHandler(NetOffice.WordApi.Range range, string name, string category, string blockType, string template);
#pragma warning restore

    #endregion

    /// <summary>
	/// CoClass Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822963.aspx </remarks>
	[SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.WordApi.EventContracts.DocumentEvents), typeof(NetOffice.WordApi.EventContracts.DocumentEvents2))]
    public interface Document : _Document, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837882.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Document_NewEventHandler NewEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821870.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Document_OpenEventHandler OpenEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821142.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event Document_CloseEventHandler CloseEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838305.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Document_SyncEventHandler SyncEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197579.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Document_XMLAfterInsertEventHandler XMLAfterInsertEvent;

        /// <summary>
        /// SupportByVersion Word 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191971.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        event Document_XMLBeforeDeleteEventHandler XMLBeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834876.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlAfterAddEventHandler ContentControlAfterAddEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835805.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlBeforeDeleteEventHandler ContentControlBeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191963.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlOnExitEventHandler ContentControlOnExitEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196332.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlOnEnterEventHandler ContentControlOnEnterEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835822.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlBeforeStoreUpdateEventHandler ContentControlBeforeStoreUpdateEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192622.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_ContentControlBeforeContentUpdateEventHandler ContentControlBeforeContentUpdateEvent;

        /// <summary>
        /// SupportByVersion Word 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197904.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        event Document_BuildingBlockInsertEventHandler BuildingBlockInsertEvent;

        #endregion
    }
}
