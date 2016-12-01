using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ADODBApi
{
	///<summary>
	/// Interface ConnectionEventsVt_Deprecated 
	/// SupportByVersion ADODB, 2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class ConnectionEventsVt_Deprecated : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ConnectionEventsVt_Deprecated);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ConnectionEventsVt_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ConnectionEventsVt_Deprecated(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 InfoMessage(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pError, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "InfoMessage", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="transactionLevel">Int32 TransactionLevel</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 BeginTransComplete(Int32 transactionLevel, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(transactionLevel, pError, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "BeginTransComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 CommitTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pError, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "CommitTransComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 RollbackTransComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pError, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "RollbackTransComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="cursorType">NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType</param>
		/// <param name="lockType">NetOffice.ADODBApi.Enums.LockTypeEnum LockType</param>
		/// <param name="options">Int32 Options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 WillExecute(string source, NetOffice.ADODBApi.Enums.CursorTypeEnum cursorType, NetOffice.ADODBApi.Enums.LockTypeEnum lockType, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, cursorType, lockType, options, adStatus, pCommand, pRecordset, pConnection);
			object returnItem = Invoker.MethodReturn(this, "WillExecute", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="recordsAffected">Int32 RecordsAffected</param>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pCommand">NetOffice.ADODBApi._Command_Deprecated pCommand</param>
		/// <param name="pRecordset">NetOffice.ADODBApi._Recordset_Deprecated pRecordset</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 ExecuteComplete(Int32 recordsAffected, NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Command_Deprecated pCommand, NetOffice.ADODBApi._Recordset_Deprecated pRecordset, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsAffected, pError, adStatus, pCommand, pRecordset, pConnection);
			object returnItem = Invoker.MethodReturn(this, "ExecuteComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="connectionString">string ConnectionString</param>
		/// <param name="userID">string UserID</param>
		/// <param name="password">string Password</param>
		/// <param name="options">Int32 Options</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 WillConnect(string connectionString, string userID, string password, Int32 options, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString, userID, password, options, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "WillConnect", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="pError">NetOffice.ADODBApi.Error pError</param>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 ConnectComplete(NetOffice.ADODBApi.Error pError, NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pError, adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "ConnectComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="adStatus">NetOffice.ADODBApi.Enums.EventStatusEnum adStatus</param>
		/// <param name="pConnection">NetOffice.ADODBApi._Connection_Deprecated pConnection</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 Disconnect(NetOffice.ADODBApi.Enums.EventStatusEnum adStatus, NetOffice.ADODBApi._Connection_Deprecated pConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(adStatus, pConnection);
			object returnItem = Invoker.MethodReturn(this, "Disconnect", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}