using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    #region Delegates

    #pragma warning disable
    public delegate void DataSourceControl_CurrentEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeExpandEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeCollapseEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeFirstPageEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforePreviousPageEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeNextPageEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeLastPageEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_DataErrorEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_DataPageCompleteEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeInitialBindEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_RecordsetSaveProgressEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_AfterDeleteEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_AfterInsertEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_AfterUpdateEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeDeleteEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeInsertEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeOverwriteEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_BeforeUpdateEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_DirtyEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_RecordExitEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_UndoEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    public delegate void DataSourceControl_FocusEventHandler(NetOffice.OWC10Api.DSCEventInfo dSCEventInfo);
    #pragma warning restore

    #endregion
    
    /// <summary>
    /// CoClass DataSourceControl 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DataSourceControlEvent))]
	[TypeId("0002E553-0000-0000-C000-000000000046")]
    public interface DataSourceControl : IDataSourceControl, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_CurrentEventHandler CurrentEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeExpandEventHandler BeforeExpandEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeCollapseEventHandler BeforeCollapseEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeFirstPageEventHandler BeforeFirstPageEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforePreviousPageEventHandler BeforePreviousPageEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeNextPageEventHandler BeforeNextPageEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeLastPageEventHandler BeforeLastPageEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_DataErrorEventHandler DataErrorEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_DataPageCompleteEventHandler DataPageCompleteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeInitialBindEventHandler BeforeInitialBindEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_RecordsetSaveProgressEventHandler RecordsetSaveProgressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_AfterDeleteEventHandler AfterDeleteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_AfterInsertEventHandler AfterInsertEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_AfterUpdateEventHandler AfterUpdateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeDeleteEventHandler BeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeInsertEventHandler BeforeInsertEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeOverwriteEventHandler BeforeOverwriteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_BeforeUpdateEventHandler BeforeUpdateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_DirtyEventHandler DirtyEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_RecordExitEventHandler RecordExitEvent;
        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_UndoEventHandler UndoEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        event DataSourceControl_FocusEventHandler FocusEvent;

        #endregion
    }
}
