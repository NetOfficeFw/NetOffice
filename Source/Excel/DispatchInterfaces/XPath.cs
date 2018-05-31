using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface XPath 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840102.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface XPath : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840823.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834983.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822616.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822150.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string Value { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197699.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlMap Map { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836528.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool Repeating { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836812.aspx </remarks>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		/// <param name="repeating">optional object repeating</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace, object repeating);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836812.aspx </remarks>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void SetValue(NetOffice.ExcelApi.XmlMap map, string xPath);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836812.aspx </remarks>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835615.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		void Clear();

		#endregion
	}
}
