using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.DAOApi;

namespace NetOffice.DAOApi.Behind
{
	/// <summary>
	/// DispatchInterface _DBEngine 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _DBEngine : _DAO, NetOffice.DAOApi._DBEngine
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
                    _contractType = typeof(NetOffice.DAOApi._DBEngine);
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
                    _type = typeof(_DBEngine);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _DBEngine() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string IniPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "IniPath");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IniPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string DefaultUser
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultUser");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultUser", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string DefaultPassword
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultPassword");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultPassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int16 LoginTimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LoginTimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LoginTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Workspaces Workspaces
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Workspaces>(this, "Workspaces", typeof(NetOffice.DAOApi.Workspaces));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Errors Errors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Errors>(this, "Errors", typeof(NetOffice.DAOApi.Errors));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string SystemDB
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SystemDB");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SystemDB", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 DefaultType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DefaultType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultType", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="action">optional object action</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Idle(object action)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Idle", action);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Idle()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Idle");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		/// <param name="options">optional object options</param>
		/// <param name="srcLocale">optional object srcLocale</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CompactDatabase(string srcName, string dstName, object dstLocale, object options, object srcLocale)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CompactDatabase", new object[]{ srcName, dstName, dstLocale, options, srcLocale });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CompactDatabase(string srcName, string dstName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CompactDatabase", srcName, dstName);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CompactDatabase(string srcName, string dstName, object dstLocale)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CompactDatabase", srcName, dstName, dstLocale);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CompactDatabase(string srcName, string dstName, object dstLocale, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CompactDatabase", srcName, dstName, dstLocale, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void RepairDatabase(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RepairDatabase", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dsn">string dsn</param>
		/// <param name="driver">string driver</param>
		/// <param name="silent">bool silent</param>
		/// <param name="attributes">string attributes</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void RegisterDatabase(string dsn, string driver, bool silent, string attributes)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RegisterDatabase", dsn, driver, silent, attributes);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Workspace _30_CreateWorkspace(string name, string userName, string password)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "_30_CreateWorkspace", typeof(NetOffice.DAOApi.Workspace), name, userName, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly, object connect)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", typeof(NetOffice.DAOApi.Database), name, options, readOnly, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database OpenDatabase(string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", typeof(NetOffice.DAOApi.Database), name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database OpenDatabase(string name, object options)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", typeof(NetOffice.DAOApi.Database), name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", typeof(NetOffice.DAOApi.Database), name, options, readOnly);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		/// <param name="option">optional object option</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database CreateDatabase(string name, string locale, object option)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", typeof(NetOffice.DAOApi.Database), name, locale, option);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Database CreateDatabase(string name, string locale)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", typeof(NetOffice.DAOApi.Database), name, locale);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void FreeLocks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FreeLocks");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void BeginTrans()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginTrans");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">optional Int32 Option = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CommitTrans(object option)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CommitTrans", option);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void CommitTrans()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CommitTrans");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void Rollback()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rollback");
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void SetDefaultWorkspace(string name, string password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultWorkspace", name, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int16 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void SetDataAccessOption(Int16 option, object value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDataAccessOption", option, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		/// <param name="reset">optional object reset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 ISAMStats(Int32 statNum, object reset)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ISAMStats", statNum, reset);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual Int32 ISAMStats(Int32 statNum)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ISAMStats", statNum);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		/// <param name="useType">optional object useType</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password, object useType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "CreateWorkspace", typeof(NetOffice.DAOApi.Workspace), name, userName, password, useType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "CreateWorkspace", typeof(NetOffice.DAOApi.Workspace), name, userName, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly, object connect)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", typeof(NetOffice.DAOApi.Connection), name, options, readOnly, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Connection OpenConnection(string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", typeof(NetOffice.DAOApi.Connection), name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Connection OpenConnection(string name, object options)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", typeof(NetOffice.DAOApi.Connection), name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", typeof(NetOffice.DAOApi.Connection), name, options, readOnly);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual void SetOption(Int32 option, object value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOption", option, value);
		}

		#endregion

		#pragma warning restore
	}
}


