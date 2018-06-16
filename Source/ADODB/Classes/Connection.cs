using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Connection_InfoMessageEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_BeginTransCompleteEventHandler(Int32 transactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_CommitTransCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_RollbackTransCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_WillExecuteEventHandler(ref string source, NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType, NetOffice.ADODBApi.Enums.LockTypeEnum lockType, ref Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command pCommand, NetOffice.ADODBApi._Recordset pRecordset, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_ExecuteCompleteEventHandler(Int32 recordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command pCommand, NetOffice.ADODBApi._Recordset pRecordset, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_WillConnectEventHandler(ref string connectionString, ref string userID, ref string password, ref Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_ConnectCompleteEventHandler(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
	public delegate void Connection_DisconnectEventHandler(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection pConnection);
    #pragma warning restore

    #endregion
   
    /// <summary>
    /// CoClass Connection 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.ConnectionEvents))]
	[TypeId("00000514-0000-0010-8000-00AA006D2EA4")]
    public interface Connection : _Connection, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_InfoMessageEventHandler InfoMessageEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_BeginTransCompleteEventHandler BeginTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_CommitTransCompleteEventHandler CommitTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_RollbackTransCompleteEventHandler RollbackTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_WillExecuteEventHandler WillExecuteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_ExecuteCompleteEventHandler ExecuteCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_WillConnectEventHandler WillConnectEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_ConnectCompleteEventHandler ConnectCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        event Connection_DisconnectEventHandler DisconnectEvent;

        #endregion
    }
}
