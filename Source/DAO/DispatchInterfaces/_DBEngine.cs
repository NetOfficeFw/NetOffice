using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _DBEngine 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _DBEngine : _DAO
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
                    _type = typeof(_DBEngine);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _DBEngine(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _DBEngine(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DBEngine(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string IniPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IniPath");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IniPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string DefaultUser
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultUser");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultUser", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string DefaultPassword
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultPassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultPassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 LoginTimeout
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LoginTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LoginTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspaces Workspaces
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Workspaces>(this, "Workspaces", NetOffice.DAOApi.Workspaces.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Errors Errors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Errors>(this, "Errors", NetOffice.DAOApi.Errors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string SystemDB
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SystemDB");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SystemDB", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 DefaultType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DefaultType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultType", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="action">optional object action</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Idle(object action)
		{
			 Factory.ExecuteMethod(this, "Idle", action);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void Idle()
		{
			 Factory.ExecuteMethod(this, "Idle");
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
		public void CompactDatabase(string srcName, string dstName, object dstLocale, object options, object srcLocale)
		{
			 Factory.ExecuteMethod(this, "CompactDatabase", new object[]{ srcName, dstName, dstLocale, options, srcLocale });
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName)
		{
			 Factory.ExecuteMethod(this, "CompactDatabase", srcName, dstName);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="srcName">string srcName</param>
		/// <param name="dstName">string dstName</param>
		/// <param name="dstLocale">optional object dstLocale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName, object dstLocale)
		{
			 Factory.ExecuteMethod(this, "CompactDatabase", srcName, dstName, dstLocale);
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
		public void CompactDatabase(string srcName, string dstName, object dstLocale, object options)
		{
			 Factory.ExecuteMethod(this, "CompactDatabase", srcName, dstName, dstLocale, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		public void RepairDatabase(string name)
		{
			 Factory.ExecuteMethod(this, "RepairDatabase", name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="dsn">string dsn</param>
		/// <param name="driver">string driver</param>
		/// <param name="silent">bool silent</param>
		/// <param name="attributes">string attributes</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void RegisterDatabase(string dsn, string driver, bool silent, string attributes)
		{
			 Factory.ExecuteMethod(this, "RegisterDatabase", dsn, driver, silent, attributes);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace _30_CreateWorkspace(string name, string userName, string password)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "_30_CreateWorkspace", NetOffice.DAOApi.Workspace.LateBindingApiWrapperType, name, userName, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly, object connect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, options, readOnly, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "OpenDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, options, readOnly);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		/// <param name="option">optional object option</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string locale, object option)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, locale, option);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="locale">string locale</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string locale)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, locale);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void FreeLocks()
		{
			 Factory.ExecuteMethod(this, "FreeLocks");
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
		/// <param name="option">optional Int32 Option = 0</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void CommitTrans(object option)
		{
			 Factory.ExecuteMethod(this, "CommitTrans", option);
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
		/// <param name="password">string password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void SetDefaultWorkspace(string name, string password)
		{
			 Factory.ExecuteMethod(this, "SetDefaultWorkspace", name, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int16 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void SetDataAccessOption(Int16 option, object value)
		{
			 Factory.ExecuteMethod(this, "SetDataAccessOption", option, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		/// <param name="reset">optional object reset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 ISAMStats(Int32 statNum, object reset)
		{
			return Factory.ExecuteInt32MethodGet(this, "ISAMStats", statNum, reset);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="statNum">Int32 statNum</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 ISAMStats(Int32 statNum)
		{
			return Factory.ExecuteInt32MethodGet(this, "ISAMStats", statNum);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		/// <param name="useType">optional object useType</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password, object useType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "CreateWorkspace", NetOffice.DAOApi.Workspace.LateBindingApiWrapperType, name, userName, password, useType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="userName">string userName</param>
		/// <param name="password">string password</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Workspace>(this, "CreateWorkspace", NetOffice.DAOApi.Workspace.LateBindingApiWrapperType, name, userName, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly, object connect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType, name, options, readOnly, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType, name, options);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="options">optional object options</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Connection>(this, "OpenConnection", NetOffice.DAOApi.Connection.LateBindingApiWrapperType, name, options, readOnly);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="value">object value</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public void SetOption(Int32 option, object value)
		{
			 Factory.ExecuteMethod(this, "SetOption", option, value);
		}

		#endregion

		#pragma warning restore
	}
}
