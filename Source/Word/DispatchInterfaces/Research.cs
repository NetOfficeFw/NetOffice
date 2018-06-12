using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Research 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194717.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Research : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192412.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196335.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845563.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840115.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string FavoriteService { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		/// <param name="launchQuery">optional bool LaunchQuery = true</param>
		[SupportByVersion("Word", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		object Query(string serviceID);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		object Query(string serviceID, object queryString);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		object Query(string serviceID, object queryString, object queryLanguage, object useSelection);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834572.aspx </remarks>
		/// <param name="languageFrom">NetOffice.WordApi.Enums.WdLanguageID languageFrom</param>
		/// <param name="languageTo">NetOffice.WordApi.Enums.WdLanguageID languageTo</param>
		[SupportByVersion("Word", 12,14,15,16)]
		object SetLanguagePair(NetOffice.WordApi.Enums.WdLanguageID languageFrom, NetOffice.WordApi.Enums.WdLanguageID languageTo);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835810.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Word", 12,14,15,16)]
		bool IsResearchService(string serviceID);

		#endregion
	}
}
