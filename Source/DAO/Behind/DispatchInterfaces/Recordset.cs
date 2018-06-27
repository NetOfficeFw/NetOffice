using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.DAOApi;

namespace NetOffice.DAOApi.Behind
{
	/// <summary>
	/// DispatchInterface Recordset 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset : _DAO, NetOffice.DAOApi.Recordset
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
                    _contractType = typeof(NetOffice.DAOApi.Recordset);
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
                    _type = typeof(Recordset);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Recordset() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool BOF
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BOF");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual byte[] Bookmark
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
		public virtual bool Bookmarkable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Bookmarkable");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual object DateCreated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DateCreated");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool EOF
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EOF");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Index");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual byte[] LastModified
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
		public virtual object LastUpdated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LastUpdated");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool LockEdits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LockEdits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LockEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool NoMatch
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NoMatch");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Sort
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Sort");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Sort", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool Transactions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Transactions");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int16 Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 RecordCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool Updatable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Updatable");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool Restartable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Restartable");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string ValidationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ValidationText");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string ValidationRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ValidationRule");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual byte[] CacheStart
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
		public virtual Int32 CacheSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CacheSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CacheSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Single PercentPosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PercentPosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PercentPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 AbsolutePosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AbsolutePosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AbsolutePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int16 EditMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 ODBCFetchCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ODBCFetchCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 ODBCFetchDelay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ODBCFetchDelay");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.DAOApi.Database Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Database>(this, "Parent", typeof(NetOffice.DAOApi.Database));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Fields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Fields>(this, "Fields", typeof(NetOffice.DAOApi.Fields));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Indexes Indexes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Indexes>(this, "Indexes", typeof(NetOffice.DAOApi.Indexes));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_Collect(object item)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Collect", item);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_Collect(object item, object value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "Collect", item, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Alias for get_Collect
		/// </summary>
		/// <param name="item">object item</param>
		[SupportByVersion("DAO", 3.6,12.0), Redirect("get_Collect")]
		public virtual object Collect(object item)
		{
			return get_Collect(item);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 hStmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hStmt");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool StillExecuting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "StillExecuting");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 BatchSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BatchSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BatchSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 BatchCollisionCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BatchCollisionCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual object BatchCollisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BatchCollisions");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Connection Connection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Connection>(this, "Connection", typeof(NetOffice.DAOApi.Connection));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int16 RecordStatus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RecordStatus");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 UpdateOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UpdateOptions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UpdateOptions", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void _30_CancelUpdate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_30_CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void AddNew()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddNew");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(object type, object options)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset OpenRecordset()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(object type)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Edit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Edit");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FindFirst(string criteria)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindFirst", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FindLast(string criteria)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindLast", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FindNext(string criteria)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindNext", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FindPrevious(string criteria)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FindPrevious", criteria);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void MoveFirst()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveFirst");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void _30_MoveLast()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_30_MoveLast");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void MoveNext()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveNext");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void MovePrevious()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MovePrevious");
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12, object key13)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12, key13 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Seek(string comparison, object key1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", comparison, key1);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="comparison">string comparison</param>
		/// <param name="key1">object key1</param>
		/// <param name="key2">optional object key2</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Seek(string comparison, object key1, object key2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", comparison, key1, key2);
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
		public virtual void Seek(string comparison, object key1, object key2, object key3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", comparison, key1, key2, key3);
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11 });
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
		public virtual void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", new object[]{ comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12 });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void _30_Update()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_30_Update");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset Clone()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "Clone");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="newQueryDef">optional object newQueryDef</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Requery(object newQueryDef)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery", newQueryDef);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Move(Int32 rows, object startBookmark)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", rows, startBookmark);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">Int32 rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Move(Int32 rows)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", rows);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		/// <param name="startBookmark">optional object startBookmark</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FillCache(object rows, object startBookmark)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FillCache", rows, startBookmark);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FillCache()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FillCache");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="rows">optional object rows</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FillCache(object rows)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FillCache", rows);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		/// <param name="inconsistent">optional object inconsistent</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset CreateDynaset(object options, object inconsistent)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", options, inconsistent);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset CreateDynaset()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset CreateDynaset(object options)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateDynaset", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset CreateSnapshot(object options)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset CreateSnapshot()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "CreateSnapshot");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.QueryDef CopyQueryDef()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CopyQueryDef", typeof(NetOffice.DAOApi.QueryDef));
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset ListFields()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListFields");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[BaseResult]
		public virtual NetOffice.DAOApi.Recordset ListIndexes()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "ListIndexes");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual object GetRows(object numRows)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows", numRows);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual object GetRows()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Cancel()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cancel");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool NextRecordset()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "NextRecordset");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CancelUpdate(object updateType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelUpdate", updateType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CancelUpdate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		/// <param name="force">optional bool Force = false</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Update(object updateType, object force)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", updateType, force);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Update()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Update(object updateType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", updateType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void MoveLast(object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLast", options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void MoveLast()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLast");
		}

		#endregion

		#pragma warning restore
	}
}


