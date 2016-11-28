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
	/// DispatchInterface Connection 
	/// SupportByVersion DAO, 3.6,12.0
	///</summary>
	[SupportByVersionAttribute("DAO", 3.6,12.0)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Connection : COMObject
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
                    _type = typeof(Connection);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Connect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connect", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Database Database
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Database", paramsArray);
				NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hDbc
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "hDbc", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int16 QueryTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QueryTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "QueryTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool Transactions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Transactions", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 RecordsAffected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordsAffected", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool StillExecuting
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StillExecuting", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool Updatable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Updatable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDefs QueryDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QueryDefs", paramsArray);
				NetOffice.DAOApi.QueryDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.QueryDefs.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordsets Recordsets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Recordsets", paramsArray);
				NetOffice.DAOApi.Recordsets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Recordsets.LateBindingApiWrapperType) as NetOffice.DAOApi.Recordsets;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Cancel()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cancel", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="sQLText">optional object SQLText</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name, object sQLText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, sQLText);
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CreateQueryDef(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="query">string Query</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Execute(string query, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(query, options);
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="query">string Query</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Execute(string query)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(query);
			Invoker.Method(this, "Execute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		/// <param name="lockEdit">optional object LockEdit</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options, object lockEdit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, options, lockEdit);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(string name, object type, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, options);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}