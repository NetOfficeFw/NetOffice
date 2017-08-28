using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IOLEDBConnection 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IOLEDBConnection : COMObject
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
                    _type = typeof(IOLEDBConnection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IOLEDBConnection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IOLEDBConnection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOLEDBConnection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object ADOConnection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ADOConnection");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool BackgroundQuery
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundQuery");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundQuery", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object CommandText
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCmdType CommandType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCmdType>(this, "CommandType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CommandType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Connection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Connection");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool EnableRefresh
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableRefresh");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableRefresh", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object LocalConnection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LocalConnection");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LocalConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool MaintainConnection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaintainConnection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaintainConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public DateTime RefreshDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "RefreshDate");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool Refreshing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Refreshing");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool RefreshOnFileOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RefreshOnFileOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshOnFileOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 RefreshPeriod
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RefreshPeriod");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshPeriod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlRobustConnect RobustConnect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlRobustConnect>(this, "RobustConnect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RobustConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool SavePassword
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SavePassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SavePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string SourceConnectionFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceConnectionFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceConnectionFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string SourceDataFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceDataFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceDataFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool OLAP
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OLAP");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool UseLocalConnection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseLocalConnection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseLocalConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 MaxDrillthroughRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxDrillthroughRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxDrillthroughRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool IsConnected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCredentialsMethod ServerCredentialsMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCredentialsMethod>(this, "ServerCredentialsMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ServerCredentialsMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string ServerSSOApplicationID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerSSOApplicationID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerSSOApplicationID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AlwaysUseConnectionFile
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlwaysUseConnectionFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlwaysUseConnectionFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ServerFillColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerFillColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFillColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ServerFontStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerFontStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFontStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ServerNumberFormat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerNumberFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerNumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ServerTextColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerTextColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool RetrieveInOfficeUILang
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RetrieveInOfficeUILang");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RetrieveInOfficeUILang", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.CalculatedMembers CalculatedMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CalculatedMembers>(this, "CalculatedMembers", NetOffice.ExcelApi.CalculatedMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 LocaleID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LocaleID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LocaleID", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 CancelRefresh()
		{
			return Factory.ExecuteInt32MethodGet(this, "CancelRefresh");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 MakeConnection()
		{
			return Factory.ExecuteInt32MethodGet(this, "MakeConnection");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 Refresh()
		{
			return Factory.ExecuteInt32MethodGet(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		/// <param name="keywords">optional object keywords</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 SaveAsODC(string oDCFileName, object description, object keywords)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAsODC", oDCFileName, description, keywords);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="oDCFileName">string oDCFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 SaveAsODC(string oDCFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAsODC", oDCFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 SaveAsODC(string oDCFileName, object description)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAsODC", oDCFileName, description);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Reconnect()
		{
			return Factory.ExecuteInt32MethodGet(this, "Reconnect");
		}

		#endregion

		#pragma warning restore
	}
}
