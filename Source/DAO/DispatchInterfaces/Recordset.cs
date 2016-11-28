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
	/// DispatchInterface Recordset 
	/// SupportByVersion DAO, 3.6,12.0
	///</summary>
	[SupportByVersionAttribute("DAO", 3.6,12.0)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Recordset : _DAO
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
                    _type = typeof(Recordset);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool BOF
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BOF", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool Bookmarkable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bookmarkable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object DateCreated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateCreated", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool EOF
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EOF", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Filter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Filter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Filter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Index", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object LastUpdated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LastUpdated", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public bool LockEdits
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockEdits", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockEdits", paramsArray);
			}
		}

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
		public bool NoMatch
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoMatch", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string Sort
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sort", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Sort", paramsArray);
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
		public Int16 Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 RecordCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
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
		public bool Restartable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Restartable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string ValidationText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ValidationText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public string ValidationRule
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ValidationRule", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
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
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 CacheSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CacheSize", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CacheSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Single PercentPosition
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PercentPosition", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PercentPosition", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 AbsolutePosition
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AbsolutePosition", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AbsolutePosition", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int16 EditMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EditMode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ODBCFetchCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ODBCFetchCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ODBCFetchDelay
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ODBCFetchDelay", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.DAOApi.Database Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Fields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.DAOApi.Fields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Fields.LateBindingApiWrapperType) as NetOffice.DAOApi.Fields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Indexes Indexes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Indexes", paramsArray);
				NetOffice.DAOApi.Indexes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Indexes.LateBindingApiWrapperType) as NetOffice.DAOApi.Indexes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Collect(object item)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			object returnItem = Invoker.PropertyGet(this, "Collect", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Collect(object item, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			Invoker.PropertySet(this, "Collect", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Alias for get_Collect
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object Collect(object item)
		{
			return get_Collect(item);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hStmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "hStmt", paramsArray);
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 BatchSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BatchSize", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BatchSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 BatchCollisionCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BatchCollisionCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object BatchCollisions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BatchCollisions", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Connection Connection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connection", paramsArray);
				NetOffice.DAOApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Connection.LateBindingApiWrapperType) as NetOffice.DAOApi.Connection;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Connection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int16 RecordStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordStatus", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public Int32 UpdateOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UpdateOptions", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UpdateOptions", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void _30_CancelUpdate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_30_CancelUpdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void AddNew()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddNew", paramsArray);
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
		/// <param name="type">optional object Type</param>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, options);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset OpenRecordset(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "OpenRecordset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Edit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Edit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="criteria">string Criteria</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FindFirst(string criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(criteria);
			Invoker.Method(this, "FindFirst", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="criteria">string Criteria</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FindLast(string criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(criteria);
			Invoker.Method(this, "FindLast", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="criteria">string Criteria</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FindNext(string criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(criteria);
			Invoker.Method(this, "FindNext", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="criteria">string Criteria</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FindPrevious(string criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(criteria);
			Invoker.Method(this, "FindPrevious", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MoveFirst()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MoveFirst", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void _30_MoveLast()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_30_MoveLast", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MoveNext()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MoveNext", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MovePrevious()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MovePrevious", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		/// <param name="key9">optional object Key9</param>
		/// <param name="key10">optional object Key10</param>
		/// <param name="key11">optional object Key11</param>
		/// <param name="key12">optional object Key12</param>
		/// <param name="key13">optional object Key13</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12, object key13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12, key13);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		/// <param name="key9">optional object Key9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		/// <param name="key9">optional object Key9</param>
		/// <param name="key10">optional object Key10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		/// <param name="key9">optional object Key9</param>
		/// <param name="key10">optional object Key10</param>
		/// <param name="key11">optional object Key11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="comparison">string Comparison</param>
		/// <param name="key1">object Key1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="key4">optional object Key4</param>
		/// <param name="key5">optional object Key5</param>
		/// <param name="key6">optional object Key6</param>
		/// <param name="key7">optional object Key7</param>
		/// <param name="key8">optional object Key8</param>
		/// <param name="key9">optional object Key9</param>
		/// <param name="key10">optional object Key10</param>
		/// <param name="key11">optional object Key11</param>
		/// <param name="key12">optional object Key12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Seek(string comparison, object key1, object key2, object key3, object key4, object key5, object key6, object key7, object key8, object key9, object key10, object key11, object key12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison, key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12);
			Invoker.Method(this, "Seek", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void _30_Update()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "_30_Update", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset Clone()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="newQueryDef">optional object NewQueryDef</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Requery(object newQueryDef)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newQueryDef);
			Invoker.Method(this, "Requery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Requery()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Requery", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="rows">Int32 Rows</param>
		/// <param name="startBookmark">optional object StartBookmark</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Move(Int32 rows, object startBookmark)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rows, startBookmark);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="rows">Int32 Rows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Move(Int32 rows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rows);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="rows">optional object Rows</param>
		/// <param name="startBookmark">optional object StartBookmark</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FillCache(object rows, object startBookmark)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rows, startBookmark);
			Invoker.Method(this, "FillCache", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FillCache()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FillCache", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="rows">optional object Rows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void FillCache(object rows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rows);
			Invoker.Method(this, "FillCache", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="options">optional object Options</param>
		/// <param name="inconsistent">optional object Inconsistent</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(object options, object inconsistent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options, inconsistent);
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="options">optional object Options</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateDynaset(object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options);
			object returnItem = Invoker.MethodReturn(this, "CreateDynaset", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="options">optional object Options</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot(object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options);
			object returnItem = Invoker.MethodReturn(this, "CreateSnapshot", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset CreateSnapshot()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateSnapshot", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.QueryDef CopyQueryDef()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CopyQueryDef", paramsArray);
			NetOffice.DAOApi.QueryDef newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.QueryDef.LateBindingApiWrapperType) as NetOffice.DAOApi.QueryDef;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset ListFields()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ListFields", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Recordset ListIndexes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ListIndexes", paramsArray);
			NetOffice.DAOApi.Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.DAOApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="numRows">optional object NumRows</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object GetRows(object numRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows);
			object returnItem = Invoker.MethodReturn(this, "GetRows", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public object GetRows()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetRows", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

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
		public bool NextRecordset()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CancelUpdate(object updateType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(updateType);
			Invoker.Method(this, "CancelUpdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void CancelUpdate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CancelUpdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		/// <param name="force">optional bool Force = false</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Update(object updateType, object force)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(updateType, force);
			Invoker.Method(this, "Update", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Update()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Update", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="updateType">optional Int32 UpdateType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void Update(object updateType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(updateType);
			Invoker.Method(this, "Update", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		/// <param name="options">optional Int32 Options = 0</param>
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MoveLast(object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(options);
			Invoker.Method(this, "MoveLast", paramsArray);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 3.6,12.0)]
		public void MoveLast()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MoveLast", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}