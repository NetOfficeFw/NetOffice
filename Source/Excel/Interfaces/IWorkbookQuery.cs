using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IWorkbookQuery 
	/// SupportByVersion Excel, 16
	/// </summary>
	[SupportByVersion("Excel", 16)]
	[EntityType(EntityType.IsInterface)]
 	public class IWorkbookQuery : COMObject
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
					_type = typeof(IWorkbookQuery);
				return _type;
			}
		}
		
		#endregion
		
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IWorkbookQuery(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		public IWorkbookQuery(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
		///<param name="comProxy">inner wrapped COM proxy</param>
		///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IWorkbookQuery(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 16)]
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
		/// SupportByVersion Excel 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public string Formula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public string Description
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		[SupportByVersion("Excel", 16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="url">string url</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult Import(string url, object overwrite)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "Import", url, overwrite);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="url">string url</param>
		[CustomMethod]
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult Import(string url)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "Import", url);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="xmlData">string xmlData</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult ImportXml(string xmlData, object overwrite)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "ImportXml", xmlData, overwrite);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="xmlData">string xmlData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult ImportXml(string xmlData)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "ImportXml", xmlData);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="url">string url</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlExportResult Export(string url, object overwrite)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlExportResult>(this, "Export", url, overwrite);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="url">string url</param>
		[CustomMethod]
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlExportResult Export(string url)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ExcelApi.Enums.XlXmlExportResult>(this, "Export", url);
		}

		/// <summary>
		/// SupportByVersion Excel 16
		/// </summary>
		/// <param name="data">string data</param>
		[SupportByVersion("Excel", 16)]
		public NetOffice.ExcelApi.Enums.XlXmlExportResult ExportXml(out string data)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			data = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(data);
			object returnItem = Invoker.MethodReturn(this, "ExportXml", paramsArray, modifiers);
			data = paramsArray[0] as string;
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlExportResult)intReturnItem;
		}

		#endregion

		#pragma warning restore
	}
}
