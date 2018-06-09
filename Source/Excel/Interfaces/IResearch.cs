using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IResearch 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("000244AC-0001-0000-C000-000000000046")]
	public interface IResearch : ICOMObject
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional object useSelection</param>
		/// <param name="launchQuery">optional object launchQuery</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional object useSelection</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage, object useSelection);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		bool IsResearchService(string serviceID);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="languageFrom">Int32 languageFrom</param>
		/// <param name="languageTo">Int32 languageTo</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		object SetLanguagePair(Int32 languageFrom, Int32 languageTo);

		#endregion
	}
}
