using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Template 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194984.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Template : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837468.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835982.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821080.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197916.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835117.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.AutoTextEntries AutoTextEntries { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835800.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID LanguageID { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193370.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Saved { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835463.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdTemplateType Type { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196412.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string FullName { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195053.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object BuiltInDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836646.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object CustomDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822979.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ListTemplates ListTemplates { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839081.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839329.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.VBIDEApi.VBProject VBProject { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195005.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool KerningByAlgorithm { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823270.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdJustificationMode JustificationMode { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195318.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFarEastLineBreakLevel FarEastLineBreakLevel { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194911.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string NoLineBreakBefore { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840243.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string NoLineBreakAfter { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192157.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 NoProofing { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834292.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID FarEastLineBreakLanguage { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195119.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlockEntries BuildingBlockEntries { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834280.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlockTypes BuildingBlockTypes { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194212.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Document OpenAsDocument();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193410.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Save();

		#endregion
	}
}
