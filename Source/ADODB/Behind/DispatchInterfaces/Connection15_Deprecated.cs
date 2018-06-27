using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Connection15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Connection15_Deprecated : _ADO, NetOffice.ADODBApi.Connection15_Deprecated
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
                    _contractType = typeof(NetOffice.ADODBApi.Connection15_Deprecated);
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
                    _type = typeof(Connection15_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Connection15_Deprecated() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string ConnectionString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 CommandTimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CommandTimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 ConnectionTimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ConnectionTimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectionTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Errors Errors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Errors>(this, "Errors", typeof(NetOffice.ADODBApi.Errors));
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string DefaultDatabase
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultDatabase");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultDatabase", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.IsolationLevelEnum IsolationLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.IsolationLevelEnum>(this, "IsolationLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "IsolationLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 Attributes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Attributes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Attributes", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorLocationEnum>(this, "CursorLocation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CursorLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ConnectModeEnum>(this, "Mode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Mode", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string Provider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Provider");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Provider", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "State");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected, object options)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", typeof(NetOffice.ADODBApi._Recordset_Deprecated), commandText, recordsAffected, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="commandText">string commandText</param>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated Execute(string commandText, object recordsAffected)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", typeof(NetOffice.ADODBApi._Recordset_Deprecated), commandText, recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 BeginTrans()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void CommitTrans()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CommitTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void RollbackTrans()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RollbackTrans");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object connectionString, object userID, object password, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", connectionString, userID, password, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object connectionString)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", connectionString);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object connectionString, object userID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", connectionString, userID);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="connectionString">optional string ConnectionString = </param>
		/// <param name="userID">optional string UserID = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object connectionString, object userID, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", connectionString, userID, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		/// <param name="schemaID">optional object schemaID</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions, object schemaID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", typeof(NetOffice.ADODBApi._Recordset_Deprecated), schema, restrictions, schemaID);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", typeof(NetOffice.ADODBApi._Recordset_Deprecated), schema);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="schema">NetOffice.ADODBApi.Enums.SchemaEnum schema</param>
		/// <param name="restrictions">optional object restrictions</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated OpenSchema(NetOffice.ADODBApi.Enums.SchemaEnum schema, object restrictions)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "OpenSchema", typeof(NetOffice.ADODBApi._Recordset_Deprecated), schema, restrictions);
		}

		#endregion

		#pragma warning restore
	}
}


