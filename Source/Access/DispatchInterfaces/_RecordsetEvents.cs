using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _RecordsetEvents 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("45165490-EF32-11D0-86FB-006097C9818C")]
	public interface _RecordsetEvents : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FieldChangeComplete(Int32 cFields, object fields, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fMoreData">Int16 fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void EndOfRecordset(Int16 fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void FetchComplete(object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset);

		#endregion
	}
}
