using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Find 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839118.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000209B0-0000-0000-C000-000000000046")]
	public interface Find : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196396.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839624.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834556.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839325.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Forward { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822678.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Font Font { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838143.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Found { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845697.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchAllWordForms { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837923.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchCase { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838695.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchWildcards { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821942.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchSoundsLike { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835745.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchWholeWord { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821682.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchFuzzy { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838094.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchByte { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836406.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ParagraphFormat ParagraphFormat { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		object Style { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string Text { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID LanguageID { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821028.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Highlight { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836618.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Replacement Replacement { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197498.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Frame Frame { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdFindWrap Wrap { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834863.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Format { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836860.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821910.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool CorrectHangulEndings { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195417.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 NoProofing { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845200.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchKashida { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839133.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchDiacritics { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845597.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchAlefHamza { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool MatchControl { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191768.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool MatchPhrase { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197820.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool MatchPrefix { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839710.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool MatchSuffix { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821316.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool IgnoreSpace { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194518.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool IgnorePunct { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835442.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool HanjaPhoneticHangul { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834930.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ClearFormatting();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194281.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SetAllFuzzyOptions();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ClearAllFuzzyOptions();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		/// <param name="hanjaPhoneticHangul">optional object hanjaPhoneticHangul</param>
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct, object hanjaPhoneticHangul);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834830.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ClearHitHighlight();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace, object ignorePunct);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007();

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace);

		#endregion
	}
}
