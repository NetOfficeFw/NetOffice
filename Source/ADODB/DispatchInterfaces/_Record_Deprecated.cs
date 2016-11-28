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
	/// DispatchInterface _Record_Deprecated 
	/// SupportByVersion ADODB, 2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Record_Deprecated : _ADO
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
                    _type = typeof(_Record_Deprecated);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Record_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Record_Deprecated(string progId) : base(progId)
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
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		public NetOffice.ADODBApi.Fields_Deprecated Fields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fields", paramsArray);
				NetOffice.ADODBApi.Fields_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Fields_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi.Fields_Deprecated;
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
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password, object options, object async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options, async);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options);
			object returnItem = Invoker.MethodReturn(this, "MoveRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password, object options, object async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options, async);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, userName, password, options);
			object returnItem = Invoker.MethodReturn(this, "CopyRecord", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void DeleteRecord(object source, object async)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, async);
			Invoker.Method(this, "DeleteRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void DeleteRecord()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeleteRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void DeleteRecord(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "DeleteRecord", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options, userName, password);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="activeConnection">optional object ActiveConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, activeConnection, mode, createOptions, options, userName);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated GetChildren()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetChildren", paramsArray);
			NetOffice.ADODBApi._Recordset_Deprecated newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType) as NetOffice.ADODBApi._Recordset_Deprecated;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
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