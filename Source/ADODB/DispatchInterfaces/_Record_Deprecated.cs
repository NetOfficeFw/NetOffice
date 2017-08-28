using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Record_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _Record_Deprecated : _ADO
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
                    _type = typeof(_Record_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Record_Deprecated(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("ADODB", 2.5)]
		public object ActiveConnection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActiveConnection");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ObjectStateEnum State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ObjectStateEnum>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public object Source
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Source");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Source", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ConnectModeEnum>(this, "Mode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Mode", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string ParentURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ParentURL");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Fields_Deprecated Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Fields_Deprecated>(this, "Fields", NetOffice.ADODBApi.Fields_Deprecated.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.RecordTypeEnum RecordType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.RecordTypeEnum>(this, "RecordType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password, object options, object async)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", new object[]{ source, destination, userName, password, options, async });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord()
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source, object destination)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", source, destination);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", source, destination, userName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", source, destination, userName, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string MoveRecord(object source, object destination, object userName, object password, object options)
		{
			return Factory.ExecuteStringMethodGet(this, "MoveRecord", new object[]{ source, destination, userName, password, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password, object options, object async)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", new object[]{ source, destination, userName, password, options, async });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord()
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source, object destination)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", source, destination);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", source, destination, userName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", source, destination, userName, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string CopyRecord(object source, object destination, object userName, object password, object options)
		{
			return Factory.ExecuteStringMethodGet(this, "CopyRecord", new object[]{ source, destination, userName, password, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="async">optional bool Async = false</param>
		[SupportByVersion("ADODB", 2.5)]
		public void DeleteRecord(object source, object async)
		{
			 Factory.ExecuteMethod(this, "DeleteRecord", source, async);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void DeleteRecord()
		{
			 Factory.ExecuteMethod(this, "DeleteRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void DeleteRecord(object source)
		{
			 Factory.ExecuteMethod(this, "DeleteRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName, object password)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options, userName, password });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source)
		{
			 Factory.ExecuteMethod(this, "Open", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection, mode);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection, mode, createOptions);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options, userName });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset_Deprecated GetChildren()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "GetChildren", NetOffice.ADODBApi._Recordset_Deprecated.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}

		#endregion

		#pragma warning restore
	}
}
