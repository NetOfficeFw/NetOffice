using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IWorkbookConnection 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	public interface IWorkbookConnection : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		string Description { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		string _Default { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlConnectionType Type { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.OLEDBConnection OLEDBConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.ODBCConnection ODBCConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Ranges Ranges { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.ModelConnection ModelConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.WorksheetDataConnection WorksheetDataConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		bool RefreshWithRefreshAll { get; set; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.TextConnection TextConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.DataFeedConnection DataFeedConnection { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		bool InModel { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.ModelTables ModelTables { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Delete();

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Refresh();

		#endregion
	}
}
