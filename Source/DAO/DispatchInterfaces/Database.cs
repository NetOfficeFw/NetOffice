using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Database 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Database : _DAO
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
                    _type = typeof(Database);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Database(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Database(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Database(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 CollatingOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CollatingOrder");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Connect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Connect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Connect", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 QueryTimeout
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "QueryTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "QueryTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Transactions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Transactions");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Updatable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Updatable");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 RecordsAffected
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecordsAffected");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDefs TableDefs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.TableDefs>(this, "TableDefs", NetOffice.DAOApi.TableDefs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDefs QueryDefs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.QueryDefs>(this, "QueryDefs", NetOffice.DAOApi.QueryDefs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relations Relations
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Relations>(this, "Relations", NetOffice.DAOApi.Relations.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Containers Containers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Containers>(this, "Containers", NetOffice.DAOApi.Containers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordsets Recordsets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Recordsets>(this, "Recordsets", NetOffice.DAOApi.Recordsets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string ReplicaID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReplicaID");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string DesignMasterID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DesignMasterID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DesignMasterID", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection Connection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Connection>(this, "Connection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Execute(string query, object options)
		{
			 Factory.ExecuteMethod(this, "Execute", query, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Execute(string query)
		{
			 Factory.ExecuteMethod(this, "Execute", query);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset", name, type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset", name, type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		/// <param name="dDL">optional object dDL</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value, object dDL)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type, value, dDL);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		/// <param name="foreignTable">optional object foreignTable</param>
		/// <param name="attributes">optional object attributes</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable, object attributes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Relation>(this, "CreateRelation", NetOffice.DAOApi.Relation.LateBindingApiWrapperType, name, table, foreignTable, attributes);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Relation>(this, "CreateRelation", NetOffice.DAOApi.Relation.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Relation>(this, "CreateRelation", NetOffice.DAOApi.Relation.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Relation>(this, "CreateRelation", NetOffice.DAOApi.Relation.LateBindingApiWrapperType, name, table);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="table">optional object table</param>
		/// <param name="foreignTable">optional object foreignTable</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Relation>(this, "CreateRelation", NetOffice.DAOApi.Relation.LateBindingApiWrapperType, name, table, foreignTable);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		/// <param name="sourceTableName">optional object sourceTableName</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName, object connect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.TableDef>(this, "CreateTableDef", NetOffice.DAOApi.TableDef.LateBindingApiWrapperType, name, attributes, sourceTableName, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.TableDef>(this, "CreateTableDef", NetOffice.DAOApi.TableDef.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.TableDef>(this, "CreateTableDef", NetOffice.DAOApi.TableDef.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.TableDef>(this, "CreateTableDef", NetOffice.DAOApi.TableDef.LateBindingApiWrapperType, name, attributes);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="attributes">optional object attributes</param>
		/// <param name="sourceTableName">optional object sourceTableName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.TableDef>(this, "CreateTableDef", NetOffice.DAOApi.TableDef.LateBindingApiWrapperType, name, attributes, sourceTableName);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void BeginTrans()
		{
			 Factory.ExecuteMethod(this, "BeginTrans");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CommitTrans(object options)
		{
			 Factory.ExecuteMethod(this, "CommitTrans", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CommitTrans()
		{
			 Factory.ExecuteMethod(this, "CommitTrans");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Rollback()
		{
			 Factory.ExecuteMethod(this, "Rollback");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="inconsistent">optional object inconsistent</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name, object options, object inconsistent)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", name, options, inconsistent);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="sQLText">optional object sQLText</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name, object sQLText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType, name, sQLText);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="source">string source</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset CreateSnapshot(string source, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot", source, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="source">string source</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot(string source)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot", source);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void DeleteQueryDef(string name)
		{
			 Factory.ExecuteMethod(this, "DeleteQueryDef", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="sQL">string sQL</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 ExecuteSQL(string sQL)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExecuteSQL", sQL);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset ListFields(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListFields", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset ListTables()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListTables");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef OpenQueryDef(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "OpenQueryDef", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset OpenTable(string name, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenTable", name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenTable(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenTable", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		/// <param name="exchangeType">optional object exchangeType</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Synchronize(string dbPathName, object exchangeType)
		{
			 Factory.ExecuteMethod(this, "Synchronize", dbPathName, exchangeType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Synchronize(string dbPathName)
		{
			 Factory.ExecuteMethod(this, "Synchronize", dbPathName);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="pathName">string pathName</param>
		/// <param name="description">string description</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MakeReplica(string pathName, string description, object options)
		{
			 Factory.ExecuteMethod(this, "MakeReplica", pathName, description, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="pathName">string pathName</param>
		/// <param name="description">string description</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MakeReplica(string pathName, string description)
		{
			 Factory.ExecuteMethod(this, "MakeReplica", pathName, description);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="bstrOld">string bstrOld</param>
		/// <param name="bstrNew">string bstrNew</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void NewPassword(string bstrOld, string bstrNew)
		{
			 Factory.ExecuteMethod(this, "NewPassword", bstrOld, bstrNew);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		/// <param name="lockEdit">optional object lockEdit</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options, object lockEdit)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type, options, lockEdit);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dbPathName">string dbPathName</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void PopulatePartial(string dbPathName)
		{
			 Factory.ExecuteMethod(this, "PopulatePartial", dbPathName);
		}

		#endregion

		#pragma warning restore
	}
}
