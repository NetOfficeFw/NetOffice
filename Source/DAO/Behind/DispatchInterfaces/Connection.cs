using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.DAOApi;

namespace NetOffice.DAOApi.Behind
{
	/// <summary>
	/// DispatchInterface Connection 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Connection : COMObject, NetOffice.DAOApi.Connection
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
                    _contractType = typeof(NetOffice.DAOApi.Connection);
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
                    _type = typeof(Connection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Connection() : base()
		{

		}

		#endregion
		
		#region Properties

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
		public virtual string Connect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Connect");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database Database
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Database>(this, "Database", typeof(NetOffice.DAOApi.Database));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 hDbc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hDbc");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int16 QueryTimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "QueryTimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "QueryTimeout", value);
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
		public virtual Int32 RecordsAffected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecordsAffected");
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
		public virtual NetOffice.DAOApi.QueryDefs QueryDefs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.QueryDefs>(this, "QueryDefs", typeof(NetOffice.DAOApi.QueryDefs));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordsets Recordsets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Recordsets>(this, "Recordsets", typeof(NetOffice.DAOApi.Recordsets));
			}
		}

		#endregion

		#region Methods

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
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="sQLText">optional object sQLText</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.QueryDef CreateQueryDef(object name, object sQLText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", typeof(NetOffice.DAOApi.QueryDef), name, sQLText);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.QueryDef CreateQueryDef()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", typeof(NetOffice.DAOApi.QueryDef));
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.QueryDef CreateQueryDef(object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.QueryDef>(this, "CreateQueryDef", typeof(NetOffice.DAOApi.QueryDef), name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		/// <param name="options">optional object options</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Execute(string query, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", query, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="query">string query</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Execute(string query)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute", query);
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
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options, object lockEdit)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type, options, lockEdit);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(string name, object type)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type);
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
		public virtual NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.DAOApi.Recordset>(this, "OpenRecordset", name, type, options);
		}

		#endregion

		#pragma warning restore
	}
}


