using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Recordset 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset : _DAO
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
                    _type = typeof(Recordset);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Recordset(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Recordset(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Recordset(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool BOF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BOF");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public byte[] Bookmark
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "Bookmark", paramsArray);
				return (byte[])returnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Bookmark", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Bookmarkable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Bookmarkable");
			}
		}

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
		public bool EOF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EOF");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Index
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Index");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public byte[] LastModified
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "LastModified", paramsArray);
				return (byte[])returnItem;
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
		public bool LockEdits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockEdits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockEdits", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool NoMatch
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NoMatch");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string Sort
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Sort");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Sort", value);
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
		public Int16 Type
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 RecordCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecordCount");
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
		public bool Restartable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Restartable");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string ValidationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ValidationText");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string ValidationRule
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ValidationRule");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public byte[] CacheStart
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "CacheStart", paramsArray);
				return (byte[])returnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CacheStart", paramsArray);
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
		public Single PercentPosition
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PercentPosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PercentPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 AbsolutePosition
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AbsolutePosition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AbsolutePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 EditMode
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ODBCFetchCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ODBCFetchCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ODBCFetchDelay
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ODBCFetchDelay");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.DAOApi.Database Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Database>(this, "Parent", NetOffice.DAOApi.Database.LateBindingApiWrapperType);
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
		public NetOffice.DAOApi.Indexes Indexes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Indexes>(this, "Indexes", NetOffice.DAOApi.Indexes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Collect(object item)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Collect", item);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Collect(object item, object value)
		{
			Factory.ExecutePropertySet(this, "Collect", item, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Alias for get_Collect
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0), Redirect("get_Collect")]
		public object Collect(object item)
		{
			return get_Collect(item);
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
		public Int32 BatchSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BatchSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BatchSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 BatchCollisionCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BatchCollisionCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object BatchCollisions
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BatchCollisions");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection Connection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Connection>(this, "Connection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 RecordStatus
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RecordStatus");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 UpdateOptions
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UpdateOptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateOptions", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void _30_CancelUpdate()
		{
			 Factory.ExecuteMethod(this, "_30_CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void AddNew()
		{
			 Factory.ExecuteMethod(this, "AddNew");
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
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type, options);
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
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Edit()
		{
			 Factory.ExecuteMethod(this, "Edit");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FindFirst(string criteria)
		{
			 Factory.ExecuteMethod(this, "FindFirst", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FindLast(string criteria)
		{
			 Factory.ExecuteMethod(this, "FindLast", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FindNext(string criteria)
		{
			 Factory.ExecuteMethod(this, "FindNext", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FindPrevious(string criteria)
		{
			 Factory.ExecuteMethod(this, "FindPrevious", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MoveFirst()
		{
			 Factory.ExecuteMethod(this, "MoveFirst");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void _30_MoveLast()
		{
			 Factory.ExecuteMethod(this, "_30_MoveLast");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MoveNext()
		{
			 Factory.ExecuteMethod(this, "MoveNext");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MovePrevious()
		{
			 Factory.ExecuteMethod(this, "MovePrevious");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		/// <param name="key12">optional object key12</param>
		/// <param name="key13">optional object key13</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12, object key13)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12, key13 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1)
		{
			 Factory.ExecuteMethod(this, "Seek", comparison, key1);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2)
		{
			 Factory.ExecuteMethod(this, "Seek", comparison, key1, key2);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3)
		{
			 Factory.ExecuteMethod(this, "Seek", comparison, key1, key2, key3);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="key4">optional object key4</param>
		/// <param name="key5">optional object key5</param>
		/// <param name="key6">optional object key6</param>
		/// <param name="key7">optional object key7</param>
		/// <param name="key8">optional object key8</param>
		/// <param name="key9">optional object key9</param>
		/// <param name="key10">optional object key10</param>
		/// <param name="key11">optional object key11</param>
		/// <param name="key12">optional object key12</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12)
		{
			 Factory.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void _30_Update()
		{
			 Factory.ExecuteMethod(this, "_30_Update");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset Clone()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "Clone");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="newQueryDef">optional object newQueryDef</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Requery(object newQueryDef)
		{
			 Factory.ExecuteMethod(this, "Requery", newQueryDef);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Requery()
		{
			 Factory.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Move(Int32 rows, object startBookmark)
		{
			 Factory.ExecuteMethod(this, "Move", rows, startBookmark);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Move(Int32 rows)
		{
			 Factory.ExecuteMethod(this, "Move", rows);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FillCache(object rows, object startBookmark)
		{
			 Factory.ExecuteMethod(this, "FillCache", rows, startBookmark);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FillCache()
		{
			 Factory.ExecuteMethod(this, "FillCache");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FillCache(object rows)
		{
			 Factory.ExecuteMethod(this, "FillCache", rows);
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
		public NetOffice.DAOApi.QueryDef CopyQueryDef()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CopyQueryDef", NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset ListFields()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListFields");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public NetOffice.DAOApi.Recordset ListIndexes()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListIndexes");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object GetRows(object numRows)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows", numRows);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public object GetRows()
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows");
		}

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
		public bool NextRecordset()
		{
			return Factory.ExecuteBoolMethodGet(this, "NextRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CancelUpdate(object updateType)
		{
			 Factory.ExecuteMethod(this, "CancelUpdate", updateType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CancelUpdate()
		{
			 Factory.ExecuteMethod(this, "CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		/// <param name="force">optional bool Force = false</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Update(object updateType, object force)
		{
			 Factory.ExecuteMethod(this, "Update", updateType, force);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Update(object updateType)
		{
			 Factory.ExecuteMethod(this, "Update", updateType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MoveLast(object options)
		{
			 Factory.ExecuteMethod(this, "MoveLast", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void MoveLast()
		{
			 Factory.ExecuteMethod(this, "MoveLast");
		}

		#endregion

		#pragma warning restore
	}
}
