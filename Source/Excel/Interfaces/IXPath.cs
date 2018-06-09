using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IXPath 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("0002447E-0001-0000-C000-000000000046")]
	public interface IXPath : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
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
		[SupportByVersion("Excel", 11,12,14,15,16)]
		string Value { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.XmlMap Map { get; }

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		bool Repeating { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		/// <param name="repeating">optional object repeating</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace, object repeating);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespace">optional object selectionNamespace</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 SetValue(NetOffice.ExcelApi.XmlMap map, string xPath, object selectionNamespace);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		Int32 Clear();

		#endregion
	}
}
