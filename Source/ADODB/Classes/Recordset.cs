using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Recordset_WillChangeFieldEventHandler(Int32 cFields, object Fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FieldChangeCompleteEventHandler(Int32 cFields, object fields, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillChangeRecordEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_RecordChangeCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillChangeRecordsetEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_RecordsetChangeCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_WillMoveEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_MoveCompleteEventHandler(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_EndOfRecordsetEventHandler(ref bool fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FetchProgressEventHandler(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
	public delegate void Recordset_FetchCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset);
    #pragma warning restore

    #endregion
   
    /// <summary>
    /// CoClass Recordset 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.RecordsetEvents))]
	[TypeId("00000535-0000-0010-8000-00AA006D2EA4")]
    public interface Recordset : _Recordset, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_WillChangeFieldEventHandler WillChangeFieldEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_FieldChangeCompleteEventHandler FieldChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_WillChangeRecordEventHandler WillChangeRecordEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_RecordChangeCompleteEventHandler RecordChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_WillChangeRecordsetEventHandler WillChangeRecordsetEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_RecordsetChangeCompleteEventHandler RecordsetChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_WillMoveEventHandler WillMoveEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_MoveCompleteEventHandler MoveCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_EndOfRecordsetEventHandler EndOfRecordsetEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_FetchProgressEventHandler FetchProgressEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Recordset_FetchCompleteEventHandler FetchCompleteEvent;

        #endregion
    }
}
