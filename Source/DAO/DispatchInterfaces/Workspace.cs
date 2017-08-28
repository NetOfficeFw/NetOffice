using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Workspace 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Workspace : _DAO
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
                    _type = typeof(Workspace);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Workspace(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Workspace(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workspace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string UserName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserName");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string _30_UserName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_30_UserName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_30_UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public string _30_Password
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_30_Password");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_30_Password", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int16 IsolateODBCTrans
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsolateODBCTrans");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsolateODBCTrans", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Databases Databases
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Databases>(this, "Databases", NetOffice.DAOApi.Databases.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Users Users
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Users>(this, "Users", NetOffice.DAOApi.Users.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Groups Groups
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Groups>(this, "Groups", NetOffice.DAOApi.Groups.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 LoginTimeout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LoginTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LoginTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 DefaultCursorDriver
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DefaultCursorDriver");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultCursorDriver", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hEnv
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hEnv");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 Type
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connections Connections
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Connections>(this, "Connections", NetOffice.DAOApi.Connections.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

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
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
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
		/// <param name="connect">string connect</param>
		/// <param name="option">optional object option</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string connect, object option)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, connect, option);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="connect">string connect</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string connect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Database>(this, "CreateDatabase", NetOffice.DAOApi.Database.LateBindingApiWrapperType, name, connect);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.User CreateUser(object name, object pID, object password)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.User>(this, "CreateUser", NetOffice.DAOApi.User.LateBindingApiWrapperType, name, pID, password);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.User CreateUser()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.User>(this, "CreateUser", NetOffice.DAOApi.User.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.User CreateUser(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.User>(this, "CreateUser", NetOffice.DAOApi.User.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.User CreateUser(object name, object pID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.User>(this, "CreateUser", NetOffice.DAOApi.User.LateBindingApiWrapperType, name, pID);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="pID">optional object pID</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Group CreateGroup(object name, object pID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Group>(this, "CreateGroup", NetOffice.DAOApi.Group.LateBindingApiWrapperType, name, pID);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Group CreateGroup()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Group>(this, "CreateGroup", NetOffice.DAOApi.Group.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Group CreateGroup(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Group>(this, "CreateGroup", NetOffice.DAOApi.Group.LateBindingApiWrapperType, name);
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

		#endregion

		#pragma warning restore
	}
}
