using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface OLEDBConnection 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OLEDBConnection : COMObject
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
                    _type = typeof(OLEDBConnection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OLEDBConnection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OLEDBConnection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEDBConnection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ADOConnection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.BackgroundQuery"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.CommandText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.CommandType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Connection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.EnableRefresh"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.LocalConnection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.MaintainConnection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.RefreshDate"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Refreshing"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.RefreshOnFileOpen"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.RefreshPeriod"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.RobustConnect"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SavePassword"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SourceConnectionFile"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SourceDataFile"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.OLAP"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.UseLocalConnection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.MaxDrillthroughRecords"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.IsConnected"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerCredentialsMethod"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerSSOApplicationID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.AlwaysUseConnectionFile"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerFillColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerFontStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerNumberFormat"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.ServerTextColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.RetrieveInOfficeUILang"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.CalculatedMembers"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.LocaleID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.CancelRefresh"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CancelRefresh()
		{
			 Factory.ExecuteMethod(this, "CancelRefresh");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.MakeConnection"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void MakeConnection()
		{
			 Factory.ExecuteMethod(this, "MakeConnection");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Refresh"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		/// <param name="keywords">optional object keywords</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void SaveAsODC(string oDCFileName, object description, object keywords)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName, description, keywords);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void SaveAsODC(string oDCFileName)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void SaveAsODC(string oDCFileName, object description)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName, description);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.OLEDBConnection.Reconnect"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void Reconnect()
		{
			 Factory.ExecuteMethod(this, "Reconnect");
		}

		#endregion

		#pragma warning restore
	}
}
