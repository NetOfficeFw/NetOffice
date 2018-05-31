using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Diagram 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Diagram : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNodes Nodes { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoDiagramType Type { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState AutoLayout { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Reverse { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState AutoFormat { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void Convert(NetOffice.OfficeApi.Enums.MsoDiagramType type);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void FitText();

		#endregion
	}
}
