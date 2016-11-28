using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ADODBApi
{
	///<summary>
	/// DispatchInterface Connection15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Connection15_Deprecated : _ADO
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
                    _type = typeof(Connection15_Deprecated);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Connection15_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection15_Deprecated(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string ConnectionString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectionString", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 CommandTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 ConnectionTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectionTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Errors Errors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Errors", paramsArray);
				NetOffice.ADODBApi.Errors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Errors.LateBindingApiWrapperType) as NetOffice.ADODBApi.Errors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string DefaultDatabase
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultDatabase", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultDatabase", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.IsolationLevelEnum IsolationLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsolationLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.IsolationLevelEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsolationLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 Attributes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Attributes", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Attributes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CursorLocation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.CursorLocationEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CursorLocation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.ConnectModeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Mode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string Provider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Provider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Provider", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 State
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "State", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="commandText">string CommandText</param>
		/// <param name="recordsAffected">object RecordsAffected</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(commandText, recordsAffected, options);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="commandText">string CommandText</param>
		/// <param name="recordsAffected">object RecordsAffected</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(commandText, recordsAffected);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 BeginTrans()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BeginTrans", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void CommitTrans()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CommitTrans", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void RollbackTrans()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RollbackTrans", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object connectionString, object userID, object password, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString, userID, password, options);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object connectionString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object connectionString, object userID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString, userID);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object connectionString, object userID, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString, userID, password);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum Schema</param>
		/// <param name="restrictions">optional object Restrictions</param>
		/// <param name="schemaID">optional object SchemaID</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions, object schemaID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(schema, restrictions, schemaID);
			object returnItem = Invoker.MethodReturn(this, "OpenSchema", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum Schema</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(schema);
			object returnItem = Invoker.MethodReturn(this, "OpenSchema", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum Schema</param>
		/// <param name="restrictions">optional object Restrictions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(schema, restrictions);
			object returnItem = Invoker.MethodReturn(this, "OpenSchema", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}