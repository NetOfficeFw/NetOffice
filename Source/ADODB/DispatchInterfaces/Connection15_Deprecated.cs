using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Connection15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Connection15_Deprecated : _ADO
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
                    _type = typeof(Connection15_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Connection15_Deprecated(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("ADODB", 2.5)]
		public string ConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 CommandTimeout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CommandTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 ConnectionTimeout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ConnectionTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Errors Errors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Errors>(this, "Errors", NetOffice.ADODBApi.Errors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string DefaultDatabase
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultDatabase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultDatabase", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.IsolationLevelEnum IsolationLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.IsolationLevelEnum>(this, "IsolationLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IsolationLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Attributes
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Attributes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Attributes", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorLocationEnum>(this, "CursorLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CursorLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ConnectModeEnum>(this, "Mode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Mode", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string Provider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Provider");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Provider", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 State
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "State");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected, object options)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType, commandText, recordsAffected, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType, commandText, recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 BeginTrans()
		{
			return Factory.ExecuteInt32MethodGet(this, "BeginTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void CommitTrans()
		{
			 Factory.ExecuteMethod(this, "CommitTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void RollbackTrans()
		{
			 Factory.ExecuteMethod(this, "RollbackTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object connectionString, object userID, object password, object options)
		{
			 Factory.ExecuteMethod(this, "Open", connectionString, userID, password, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object connectionString)
		{
			 Factory.ExecuteMethod(this, "Open", connectionString);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object connectionString, object userID)
		{
			 Factory.ExecuteMethod(this, "Open", connectionString, userID);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object connectionString, object userID, object password)
		{
			 Factory.ExecuteMethod(this, "Open", connectionString, userID, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		/// <param name="schemaID">optional object schemaID</param>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions, object schemaID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType, schema, restrictions, schemaID);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType, schema);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType, schema, restrictions);
		}

		#endregion

		#pragma warning restore
	}
}
