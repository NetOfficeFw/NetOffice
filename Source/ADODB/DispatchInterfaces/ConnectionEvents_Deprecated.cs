using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface ConnectionEvents_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ConnectionEvents_Deprecated : COMObject
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
                    _type = typeof(ConnectionEvents_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ConnectionEvents_Deprecated(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ConnectionEvents_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEvents_Deprecated(string progId) : base(progId)
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
		public Int32 InfoMessage(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "InfoMessage", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="transactionLevel">Int32 transactionLevel</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 BeginTransComplete(Int32 transactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "BeginTransComplete", transactionLevel, pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 CommitTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "CommitTransComplete", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 RollbackTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "RollbackTransComplete", pError, adStatus, pConnection);
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
		public Int32 WillExecute(string source, NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType, NetOffice.ADODBApi.Enums.LockTypeEnum lockType, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillExecute", new object[]{ source, cursorType, lockType, options, adStatus, pCommand, pRecordset, pConnection });
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
		public Int32 ExecuteComplete(Int32 recordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExecuteComplete", new object[]{ recordsAffected, pError, adStatus, pCommand, pRecordset, pConnection });
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
		public Int32 WillConnect(string connectionString, string userID, string password, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "WillConnect", new object[]{ connectionString, userID, password, options, adStatus, pConnection });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 ConnectComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "ConnectComplete", pError, adStatus, pConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Disconnect(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "Disconnect", adStatus, pConnection);
		}

		#endregion

		#pragma warning restore
	}
}
