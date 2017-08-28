using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface RecordsetEvents_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class RecordsetEvents_Deprecated : COMObject
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
                    _type = typeof(RecordsetEvents_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public RecordsetEvents_Deprecated(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public RecordsetEvents_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecordsetEvents_Deprecated(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 WillChangeField(Int32 cFields, object fields, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillChangeField", cFields, fields, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="cFields">Int32 cFields</param>
		/// <param name="fields">object fields</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 FieldChangeComplete(Int32 cFields, object fields, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "FieldChangeComplete", new object[]{ cFields, fields, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 WillChangeRecord(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillChangeRecord", adReason, cRecords, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="cRecords">Int32 cRecords</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 RecordChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, Int32 cRecords, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "RecordChangeComplete", new object[]{ adReason, cRecords, pError, adStatus, pRecordset });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 WillChangeRecordset(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillChangeRecordset", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 RecordsetChangeComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "RecordsetChangeComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 WillMove(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillMove", adReason, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adReason">NetOffice.ADODBApi.Enums.EventReasonEnum adReason</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 MoveComplete(NetOffice.ADODBApi.Enums.EventReasonEnum adReason, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveComplete", adReason, pError, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fMoreData">bool fMoreData</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 EndOfRecordset(bool fMoreData, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "EndOfRecordset", fMoreData, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="progress">Int32 progress</param>
		/// <param name="maxProgress">Int32 maxProgress</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 FetchProgress(Int32 progress, Int32 maxProgress, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "FetchProgress", progress, maxProgress, adStatus, pRecordset);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 FetchComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Recordset_Deprecated pRecordset)
		{
			return Factory.ExecuteInt32MethodGet(this, "FetchComplete", pError, adStatus, pRecordset);
		}

		#endregion

		#pragma warning restore
	}
}
