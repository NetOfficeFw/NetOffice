using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _RecordsetEvents 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _RecordsetEvents : COMObject, NetOffice.AccessApi._RecordsetEvents
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
                    _contractType = typeof(NetOffice.AccessApi._RecordsetEvents);
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
                    _type = typeof(_RecordsetEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _RecordsetEvents() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WillChangeField", cFields, fields, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FieldChangeComplete(Int32 cFields, object fields, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FieldChangeComplete", new object[]{ cFields, fields, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WillChangeRecord", adReason, cRecords, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RecordChangeComplete", new object[]{ adReason, cRecords, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WillChangeRecordset", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RecordsetChangeComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WillMove", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fMoreData">Int16 fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void EndOfRecordset(Int16 fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndOfRecordset", fMoreData, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FetchProgress", progress, maxProgress, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void FetchComplete(object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FetchComplete", pError, adStatus, pRecordset);
		}

		#endregion

		#pragma warning restore
	}
}

