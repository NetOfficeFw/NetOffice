using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ContentControl 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821215.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ContentControl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845327.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193736.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840017.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839788.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Range Range { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835775.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool LockContentControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822956.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool LockContents { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193642.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.XMLMapping XMLMapping { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192401.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdContentControlType Type { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194660.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControlListEntries DropdownListEntries { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194687.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlock PlaceholderText { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194027.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string Title { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845679.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string DateDisplayFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195729.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool MultiLine { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822193.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.ContentControl ParentContentControl { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197494.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool Temporary { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192748.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string ID { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191954.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ShowingPlaceholderText { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838048.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdContentControlDateStorageFormat DateStorageFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198199.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdBuildingBlockTypes BuildingBlockType { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838162.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string BuildingBlockCategory { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193973.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID DateDisplayLocale { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820775.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		object DefaultTextStyle { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196651.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Enums.WdCalendarType DateCalendarType { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195293.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194593.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		bool Checked { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227730.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdColor Color { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227646.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdContentControlAppearance Appearance { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229832.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdContentControlLevel Level { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231663.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.RepeatingSectionItemColl RepeatingSectionItems { get; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230582.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		string RepeatingSectionItemTitle { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229741.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		bool AllowInsertDeleteSection { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838347.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void Copy();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836258.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void Cut();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194849.aspx </remarks>
		/// <param name="deleteContents">optional bool DeleteContents = false</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void Delete(object deleteContents);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194849.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		/// <param name="text">optional string Text = </param>
		[SupportByVersion("Word", 12,14,15,16)]
		[KnownIssue]
		virtual void SetPlaceholderText(object buildingBlock, object range, object text);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		[KnownIssue]
		virtual void SetPlaceholderText();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		[KnownIssue]
		virtual void SetPlaceholderText(object buildingBlock);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		[KnownIssue]
		virtual void SetPlaceholderText(object buildingBlock, object range);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195895.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		void Ungroup();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197558.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersion("Word", 14,15,16)]
		void SetCheckedSymbol(Int32 characterNumber, object font);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197558.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SetCheckedSymbol(Int32 characterNumber);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836974.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersion("Word", 14,15,16)]
		void SetUncheckedSymbol(Int32 characterNumber, object font);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836974.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void SetUncheckedSymbol(Int32 characterNumber);

		#endregion
	}
}
