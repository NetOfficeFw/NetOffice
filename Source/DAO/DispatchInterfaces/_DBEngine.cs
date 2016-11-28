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
	/// DispatchInterface _DBEngine 
	/// SupportByVersion DAO, 3.6,12.0
	///</summary>
	[SupportByVersionAttribute("DAO", 3.6,12.0)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _DBEngine : _DAO
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
                    _type = typeof(_DBEngine);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string IniPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IniPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IniPath", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string DefaultUser
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultUser", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultUser", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string DefaultPassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultPassword", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultPassword", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int16 LoginTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LoginTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LoginTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspaces Workspaces
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Workspaces", paramsArray);
				NetOffice.DAOApi.Workspaces newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Workspaces.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspaces;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Errors Errors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Errors", paramsArray);
				NetOffice.DAOApi.Errors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Errors.LateBindingApiWrapperType) as NetOffice.DAOApi.Errors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string SystemDB
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SystemDB", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SystemDB", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 DefaultType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultType", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="action">optional object Action</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Idle(object action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action);
			Invoker.Method(this, "Idle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Idle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Idle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="srcName">string SrcName</param>
		/// <param name="dstName">string DstName</param>
		/// <param name="dstLocale">optional object DstLocale</param>
		/// <param name="options">optional object Options</param>
		/// <param name="srcLocale">optional object SrcLocale</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName, object dstLocale, object options, object srcLocale)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(srcName, dstName, dstLocale, options, srcLocale);
			Invoker.Method(this, "CompactDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="srcName">string SrcName</param>
		/// <param name="dstName">string DstName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(srcName, dstName);
			Invoker.Method(this, "CompactDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="srcName">string SrcName</param>
		/// <param name="dstName">string DstName</param>
		/// <param name="dstLocale">optional object DstLocale</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName, object dstLocale)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(srcName, dstName, dstLocale);
			Invoker.Method(this, "CompactDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="srcName">string SrcName</param>
		/// <param name="dstName">string DstName</param>
		/// <param name="dstLocale">optional object DstLocale</param>
		/// <param name="options">optional object Options</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CompactDatabase(string srcName, string dstName, object dstLocale, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(srcName, dstName, dstLocale, options);
			Invoker.Method(this, "CompactDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void RepairDatabase(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "RepairDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="dsn">string Dsn</param>
		/// <param name="driver">string Driver</param>
		/// <param name="silent">bool Silent</param>
		/// <param name="attributes">string Attributes</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void RegisterDatabase(string dsn, string driver, bool silent, string attributes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dsn, driver, silent, attributes);
			Invoker.Method(this, "RegisterDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace _30_CreateWorkspace(string name, string userName, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, userName, password);
			object returnItem = Invoker.MethodReturn(this, "_30_CreateWorkspace", paramsArray);
			NetOffice.DAOApi.Workspace newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Workspace.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="connect">optional object Connect</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options, readOnly, connect);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
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
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database OpenDatabase(string name, object options, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options, readOnly);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="locale">string Locale</param>
		/// <param name="option">optional object Option</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string locale, object option)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, locale, option);
			object returnItem = Invoker.MethodReturn(this, "CreateDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="locale">string Locale</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database CreateDatabase(string name, string locale)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, locale);
			object returnItem = Invoker.MethodReturn(this, "CreateDatabase", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FreeLocks()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FreeLocks", paramsArray);
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
		/// <param name="option">optional Int32 Option = 0</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CommitTrans(object option)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(option);
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
		/// <param name="password">string Password</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void SetDefaultWorkspace(string name, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, password);
			Invoker.Method(this, "SetDefaultWorkspace", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="option">Int16 Option</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void SetDataAccessOption(Int16 option, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(option, value);
			Invoker.Method(this, "SetDataAccessOption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="statNum">Int32 StatNum</param>
		/// <param name="reset">optional object Reset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 ISAMStats(Int32 statNum, object reset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(statNum, reset);
			object returnItem = Invoker.MethodReturn(this, "ISAMStats", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="statNum">Int32 StatNum</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 ISAMStats(Int32 statNum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(statNum);
			object returnItem = Invoker.MethodReturn(this, "ISAMStats", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		/// <param name="useType">optional object UseType</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password, object useType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, userName, password, useType);
			object returnItem = Invoker.MethodReturn(this, "CreateWorkspace", paramsArray);
			NetOffice.DAOApi.Workspace newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Workspace.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Workspace CreateWorkspace(string name, string userName, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, userName, password);
			object returnItem = Invoker.MethodReturn(this, "CreateWorkspace", paramsArray);
			NetOffice.DAOApi.Workspace newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Workspace.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="connect">optional object Connect</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options, readOnly, connect);
			object returnItem = Invoker.MethodReturn(this, "OpenConnection", paramsArray);
			NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenConnection", paramsArray);
			NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
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
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options);
			object returnItem = Invoker.MethodReturn(this, "OpenConnection", paramsArray);
			NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="options">optional object Options</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection OpenConnection(string name, object options, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, options, readOnly);
			object returnItem = Invoker.MethodReturn(this, "OpenConnection", paramsArray);
			NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="option">Int32 Option</param>
		/// <param name="value">object Value</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void SetOption(Int32 option, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(option, value);
			Invoker.Method(this, "SetOption", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}