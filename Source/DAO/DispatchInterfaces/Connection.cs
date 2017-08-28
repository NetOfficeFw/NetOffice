using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Connection 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Connection : COMObject
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
                    _type = typeof(Connection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Connection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Connection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Connect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Connect");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database Database
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Database>(this, "Database", NetOffice.DAOApi.Database.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hDbc
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hDbc");
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
		public bool StillExecuting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StillExecuting");
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
		public NetOffice.DAOApi.Recordsets Recordsets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Recordsets>(this, "Recordsets", NetOffice.DAOApi.Recordsets.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}

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

		#endregion

		#pragma warning restore
	}
}
