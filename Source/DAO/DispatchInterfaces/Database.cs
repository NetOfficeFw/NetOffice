using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.DAOApi
{
	///<summary>
	/// DispatchInterface Database 
	/// SupportByVersion DAO, 3.6,12.0
	///</summary>
	[SupportByVersionAttribute("DAO", 3.6,12.0)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Database : _DAO
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
                    _type = typeof(Database);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 CollatingOrder
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CollatingOrder", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Connect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connect", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Connect", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int16 QueryTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QueryTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "QueryTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool Transactions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Transactions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool Updatable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Updatable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
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
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 RecordsAffected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordsAffected", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDefs TableDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableDefs", paramsArray);
				NetOffice.DAOApi.TableDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.TableDefs.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDefs QueryDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QueryDefs", paramsArray);
				NetOffice.DAOApi.QueryDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.QueryDefs.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relations Relations
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Relations", paramsArray);
				NetOffice.DAOApi.Relations newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Relations.LateBindingApiWrapperType) as NetOffice.DAOApi.Relations;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Containers Containers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Containers", paramsArray);
				NetOffice.DAOApi.Containers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Containers.LateBindingApiWrapperType) as NetOffice.DAOApi.Containers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordsets Recordsets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Recordsets", paramsArray);
				NetOffice.DAOApi.Recordsets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Recordsets.LateBindingApiWrapperType) as NetOffice.DAOApi.Recordsets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string ReplicaID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReplicaID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string DesignMasterID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DesignMasterID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DesignMasterID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection Connection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connection", paramsArray);
				NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="query">string Query</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Execute(string query, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(query, options);
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="query">string Query</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Execute(string query)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(query);
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, options);
			object returnItem = Invoker.MethodReturn(this, "_30_OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "_30_OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(string name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			object returnItem = Invoker.MethodReturn(this, "_30_OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="value">optional object Value</param>
		/// <param name="dDL">optional object DDL</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value, object dDL)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, value, dDL);
			object returnItem = Invoker.MethodReturn(this, "CreateProperty", paramsArray);
			NetOffice.DAOApi.Property newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Property;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateProperty", paramsArray);
			NetOffice.DAOApi.Property newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Property;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateProperty", paramsArray);
			NetOffice.DAOApi.Property newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Property;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			object returnItem = Invoker.MethodReturn(this, "CreateProperty", paramsArray);
			NetOffice.DAOApi.Property newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Property;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="value">optional object Value</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, value);
			object returnItem = Invoker.MethodReturn(this, "CreateProperty", paramsArray);
			NetOffice.DAOApi.Property newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Property;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="table">optional object Table</param>
		/// <param name="foreignTable">optional object ForeignTable</param>
		/// <param name="attributes">optional object Attributes</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable, object attributes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, table, foreignTable, attributes);
			object returnItem = Invoker.MethodReturn(this, "CreateRelation", paramsArray);
			NetOffice.DAOApi.Relation newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Relation.LateBindingApiWrapperType) as NetOffice.DAOApi.Relation;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateRelation", paramsArray);
			NetOffice.DAOApi.Relation newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Relation.LateBindingApiWrapperType) as NetOffice.DAOApi.Relation;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateRelation", paramsArray);
			NetOffice.DAOApi.Relation newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Relation.LateBindingApiWrapperType) as NetOffice.DAOApi.Relation;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="table">optional object Table</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, table);
			object returnItem = Invoker.MethodReturn(this, "CreateRelation", paramsArray);
			NetOffice.DAOApi.Relation newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Relation.LateBindingApiWrapperType) as NetOffice.DAOApi.Relation;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="table">optional object Table</param>
		/// <param name="foreignTable">optional object ForeignTable</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Relation CreateRelation(object name, object table, object foreignTable)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, table, foreignTable);
			object returnItem = Invoker.MethodReturn(this, "CreateRelation", paramsArray);
			NetOffice.DAOApi.Relation newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Relation.LateBindingApiWrapperType) as NetOffice.DAOApi.Relation;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="attributes">optional object Attributes</param>
		/// <param name="sourceTableName">optional object SourceTableName</param>
		/// <param name="connect">optional object Connect</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, attributes, sourceTableName, connect);
			object returnItem = Invoker.MethodReturn(this, "CreateTableDef", paramsArray);
			NetOffice.DAOApi.TableDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.TableDef.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateTableDef", paramsArray);
			NetOffice.DAOApi.TableDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.TableDef.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateTableDef", paramsArray);
			NetOffice.DAOApi.TableDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.TableDef.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="attributes">optional object Attributes</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, attributes);
			object returnItem = Invoker.MethodReturn(this, "CreateTableDef", paramsArray);
			NetOffice.DAOApi.TableDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.TableDef.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="attributes">optional object Attributes</param>
		/// <param name="sourceTableName">optional object SourceTableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.TableDef CreateTableDef(object name, object attributes, object sourceTableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, attributes, sourceTableName);
			object returnItem = Invoker.MethodReturn(this, "CreateTableDef", paramsArray);
			NetOffice.DAOApi.TableDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.TableDef.LateBindingApiWrapperType) as NetOffice.DAOApi.TableDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void BeginTrans()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BeginTrans", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CommitTrans(object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options);
			Invoker.Method(this, "CommitTrans", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CommitTrans()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CommitTrans", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Rollback()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Rollback", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		/// <param name="inconsistent">optional object Inconsistent</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name, object options, object inconsistent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options, inconsistent);
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(string name, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options);
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="sQLText">optional object SQLText</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name, object sQLText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, sQLText);
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot(string source, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, options);
			object returnItem = Invoker.MethodReturn(this, "CreateSnapshot", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="source">string Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot(string source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "CreateSnapshot", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void DeleteQueryDef(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "DeleteQueryDef", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="sQL">string SQL</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 ExecuteSQL(string sQL)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sQL);
			object returnItem = Invoker.MethodReturn(this, "ExecuteSQL", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset ListFields(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "ListFields", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset ListTables()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ListTables", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef OpenQueryDef(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenTable(string name, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options);
			object returnItem = Invoker.MethodReturn(this, "OpenTable", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenTable(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenTable", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="dbPathName">string DbPathName</param>
		/// <param name="exchangeType">optional object ExchangeType</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Synchronize(string dbPathName, object exchangeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dbPathName, exchangeType);
			Invoker.Method(this, "Synchronize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="dbPathName">string DbPathName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Synchronize(string dbPathName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dbPathName);
			Invoker.Method(this, "Synchronize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="pathName">string PathName</param>
		/// <param name="description">string Description</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MakeReplica(string pathName, string description, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pathName, description, options);
			Invoker.Method(this, "MakeReplica", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="pathName">string PathName</param>
		/// <param name="description">string Description</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MakeReplica(string pathName, string description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pathName, description);
			Invoker.Method(this, "MakeReplica", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="bstrOld">string bstrOld</param>
		/// <param name="bstrNew">string bstrNew</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void NewPassword(string bstrOld, string bstrNew)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrOld, bstrNew);
			Invoker.Method(this, "NewPassword", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		/// <param name="lockEdit">optional object LockEdit</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options, object lockEdit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, options, lockEdit);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, options);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="dbPathName">string DbPathName</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void PopulatePartial(string dbPathName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dbPathName);
			Invoker.Method(this, "PopulatePartial", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}