using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// Interface RecordsetEventsVt_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000403-0000-0010-8000-00AA006D2EA4")]
	public interface RecordsetEventsVt_Deprecated : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 FieldChangeComplete(Int32 cFields, object fields, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fMoreData">bool fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 EndOfRecordset(bool fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		Int32 FetchComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset);

		#endregion
	}
}
