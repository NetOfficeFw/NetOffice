using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface XmlMap 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196723.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002447B-0000-0000-C000-000000000046")]
	public interface XmlMap : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837973.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837435.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838579.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string _Default { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821797.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198192.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool IsExportable { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193545.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool ShowImportExportValidationErrors { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836522.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool SaveDataSourceDefinition { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839452.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool AdjustColumnWidth { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834417.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool PreserveColumnFilter { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195992.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool PreserveNumberFormatting { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837963.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool AppendOnImport { get; set; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193360.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string RootElementName { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841104.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlNamespace RootElementNamespace { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835875.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlSchemas Schemas { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194288.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlDataBinding DataBinding { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194285.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.WorkbookConnection WorkbookConnection { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835524.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821620.aspx </remarks>
		/// <param name="url">string url</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlImportResult Import(string url, object overwrite);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821620.aspx </remarks>
		/// <param name="url">string url</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlImportResult Import(string url);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193320.aspx </remarks>
		/// <param name="xmlData">string xmlData</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlImportResult ImportXml(string xmlData, object overwrite);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193320.aspx </remarks>
		/// <param name="xmlData">string xmlData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlImportResult ImportXml(string xmlData);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194404.aspx </remarks>
		/// <param name="url">string url</param>
		/// <param name="overwrite">optional object overwrite</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlExportResult Export(string url, object overwrite);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194404.aspx </remarks>
		/// <param name="url">string url</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlExportResult Export(string url);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841274.aspx </remarks>
		/// <param name="data">string data</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlXmlExportResult ExportXml(out string data);

		#endregion
	}
}
