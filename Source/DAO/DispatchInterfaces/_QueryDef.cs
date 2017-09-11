using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _QueryDef 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _QueryDef : _DAO
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
                    _type = typeof(_QueryDef);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _QueryDef(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _QueryDef(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryDef(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object DateCreated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DateCreated");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object LastUpdated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LastUpdated");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 ODBCTimeout
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ODBCTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ODBCTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 Type
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string SQL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SQL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SQL", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool ReturnsRecords
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReturnsRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReturnsRecords", value);
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
		public NetOffice.DAOApi.Fields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Fields>(this, "Fields", NetOffice.DAOApi.Fields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Parameters Parameters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Parameters>(this, "Parameters", NetOffice.DAOApi.Parameters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hStmt
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hStmt");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 MaxRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecords", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 CacheSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CacheSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CacheSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object Prepare
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Prepare");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Prepare", value);
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
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset", type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30_OpenRecordset(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30_OpenRecordset", type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset _30__OpenRecordset(object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30__OpenRecordset", type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30__OpenRecordset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30__OpenRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _30__OpenRecordset(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_30__OpenRecordset", type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef _Copy()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "_Copy", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Execute(object options)
		{
			 Factory.ExecuteMethod(this, "Execute", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Execute()
		{
			 Factory.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="pQdef">NetOffice.DAOApi.QueryDef pQdef</param>
		/// <param name="lps">Int16 lps</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Compare(NetOffice.DAOApi.QueryDef pQdef, Int16 lps)
		{
			 Factory.ExecuteMethod(this, "Compare", pQdef, lps);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		/// <param name="inconsistent">optional object inconsistent</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset CreateDynaset(object options, object inconsistent)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", options, inconsistent);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset CreateSnapshot(object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset ListParameters()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListParameters");
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
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		/// <param name="lockEdit">optional object lockEdit</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type, object options, object lockEdit)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type, options, lockEdit);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		/// <param name="lockEdit">optional object lockEdit</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset _OpenRecordset(object type, object options, object lockEdit)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_OpenRecordset", type, options, lockEdit);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _OpenRecordset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_OpenRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _OpenRecordset(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_OpenRecordset", type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset _OpenRecordset(object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "_OpenRecordset", type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}

		#endregion

		#pragma warning restore
	}
}
