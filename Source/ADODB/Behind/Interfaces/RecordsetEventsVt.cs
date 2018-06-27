using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// Interface RecordsetEventsVt 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsInterface)]
 	public class RecordsetEventsVt : COMObject, NetOffice.ADODBApi.RecordsetEventsVt
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ADODBApi.RecordsetEventsVt);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(RecordsetEventsVt);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public RecordsetEventsVt() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillChangeField", cFields, fields, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 FieldChangeComplete(Int32 cFields, object fields, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FieldChangeComplete", new object[]{ cFields, fields, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillChangeRecord", adReason, cRecords, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RecordChangeComplete", new object[]{ adReason, cRecords, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillChangeRecordset", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RecordsetChangeComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillMove", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fMoreData">bool fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 EndOfRecordset(bool fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndOfRecordset", fMoreData, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FetchProgress", progress, maxProgress, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset pRecordset</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 FetchComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset pRecordset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FetchComplete", pError, adStatus, pRecordset);
		}

		#endregion

		#pragma warning restore
	}
}

