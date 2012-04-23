using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ADODBApi
{
	///<summary>
	/// DispatchInterface _Record 
	/// SupportByVersion ADODB, 2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Record : _ADO
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
                    _type = typeof(_Record);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public object ActiveConnection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveConnection", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ActiveConnection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ObjectStateEnum State
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "State", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.ObjectStateEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public object Source
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Source", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Source", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.ConnectModeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Mode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string ParentURL
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentURL", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Fields Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.ADODBApi.Fields newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Fields.LateBindingApiWrapperType) as NetOffice.ADODBApi.Fields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.RecordTypeEnum RecordType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.RecordTypeEnum)intReturnItem;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		/// <param name="async">bool Async</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(string source, string destination, string userName, string password, NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum options, bool async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options, async);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(string source, string destination, string userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(string source, string destination, string userName, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(string source, string destination, string userName, string password, NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		/// <param name="async">bool Async</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(string source, string destination, string userName, string password, NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum options, bool async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options, async);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(string source, string destination, string userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(string source, string destination, string userName, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(string source, string destination, string userName, string password, NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="async">bool Async</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void DeleteRecord(string source, bool async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, async);
			Invoker.Method(this, "DeleteRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="activeConnection">object ActiveConnection</param>
		/// <param name="mode">NetOffice.ADODBApi.Enums.ConnectModeEnum Mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">string UserName</param>
		/// <param name="password">string Password</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, NetOffice.ADODBApi.Enums.ConnectModeEnum mode, NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum createOptions, NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum options, string userName, string password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options, userName, password);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="activeConnection">object ActiveConnection</param>
		/// <param name="mode">NetOffice.ADODBApi.Enums.ConnectModeEnum Mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, NetOffice.ADODBApi.Enums.ConnectModeEnum mode, NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum createOptions, NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="activeConnection">object ActiveConnection</param>
		/// <param name="mode">NetOffice.ADODBApi.Enums.ConnectModeEnum Mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">string UserName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, NetOffice.ADODBApi.Enums.ConnectModeEnum mode, NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum createOptions, NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum options, string userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options, userName);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset GetChildren()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetChildren", paramsArray);
			NetOffice.ADODBApi._Recordset newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Cancel()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cancel", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}