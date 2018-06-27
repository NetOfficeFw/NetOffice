using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface ConnectionEvents_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ConnectionEvents_Deprecated : COMObject, NetOffice.ADODBApi.ConnectionEvents_Deprecated
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
                    _contractType = typeof(NetOffice.ADODBApi.ConnectionEvents_Deprecated);
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
                    _type = typeof(ConnectionEvents_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ConnectionEvents_Deprecated() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 InfoMessage(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InfoMessage", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="transactionLevel">Int32 transactionLevel</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 BeginTransComplete(Int32 transactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginTransComplete", transactionLevel, pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 CommitTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CommitTransComplete", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 RollbackTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RollbackTransComplete", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="cursorType">NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType</param>
		/// <param name="lockType">NetOffice.ADODBApi.Enums.LockTypeEnum lockType</param>
		/// <param name="options">Int32 options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 WillExecute(string source, NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType, NetOffice.ADODBApi.Enums.LockTypeEnum lockType, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillExecute", new object[]{ source, cursorType, lockType, options, adStatus, pCommand, pRecordset, pConnection });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">Int32 recordsAffected</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 ExecuteComplete(Int32 recordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExecuteComplete", new object[]{ recordsAffected, pError, adStatus, pCommand, pRecordset, pConnection });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">string connectionString</param>
		/// <param name="userID">string userID</param>
		/// <param name="password">string password</param>
		/// <param name="options">Int32 options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 WillConnect(string connectionString, string userID, string password, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "WillConnect", new object[]{ connectionString, userID, password, options, adStatus, pConnection });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 ConnectComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConnectComplete", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 Disconnect(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Disconnect", adStatus, pConnection);
		}

		#endregion

		#pragma warning restore
	}
}

