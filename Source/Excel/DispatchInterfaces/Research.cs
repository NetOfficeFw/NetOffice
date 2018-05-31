using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Research 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840013.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Research : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838183.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838032.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837577.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840468.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840468.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840468.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840468.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840468.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196527.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		bool IsResearchService(string serviceID);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193810.aspx </remarks>
		/// <param name="languageFrom">Int32 languageFrom</param>
		/// <param name="languageTo">Int32 languageTo</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		object SetLanguagePair(Int32 languageFrom, Int32 languageTo);

		#endregion
	}
}
