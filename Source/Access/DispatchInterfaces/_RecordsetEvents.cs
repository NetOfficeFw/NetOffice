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
 	public class _RecordsetEvents : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _RecordsetEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _RecordsetEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _RecordsetEvents(string progId) : base(progId)
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
		public void WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "WillChangeField", cFields, fields, adStatus, pRecordset);
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
		public void FieldChangeComplete(Int32 cFields, object fields, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "FieldChangeComplete", new object[]{ cFields, fields, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "WillChangeRecord", adReason, cRecords, adStatus, pRecordset);
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
		public void RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "RecordChangeComplete", new object[]{ adReason, cRecords, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "WillChangeRecordset", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "RecordsetChangeComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "WillMove", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "MoveComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fMoreData">Int16 fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void EndOfRecordset(Int16 fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "EndOfRecordset", fMoreData, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "FetchProgress", progress, maxProgress, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pError">object pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">object pRecordset</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void FetchComplete(object pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, object pRecordset)
		{
			 Factory.ExecuteMethod(this, "FetchComplete", pError, adStatus, pRecordset);
		}

		#endregion

		#pragma warning restore
	}
}
