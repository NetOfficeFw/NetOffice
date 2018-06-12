using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Table 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834860.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface Table : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197195.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839082.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845545.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198160.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Columns Columns { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839587.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Rows Rows { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823239.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Borders Borders { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845404.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Shading Shading { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool Uniform { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193447.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 AutoFormatType { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836124.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Tables Tables { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 NestingLevel { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool AllowPageBreaks { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool AllowAutoFit { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single PreferredWidth { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdPreferredWidthType PreferredWidthType { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single TopPadding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838742.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single BottomPadding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836311.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single LeftPadding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single RightPadding { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196121.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Single Spacing { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdTableDirection TableDirection { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string ID { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196552.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		object Style { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191959.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ApplyStyleHeadingRows { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844792.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ApplyStyleLastRow { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834832.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ApplyStyleFirstColumn { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839138.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		bool ApplyStyleLastColumn { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192619.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ApplyStyleRowBands { get; set; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839106.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		bool ApplyStyleColumnBands { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835972.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string Title { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820918.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		string Descr { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194359.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Select();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="languageID">optional object languageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object languageID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196507.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortAscending();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835818.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void SortDescending();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838712.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void UpdateAutoFormat();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ConvertToTextOld(object separator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ConvertToTextOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821612.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Cell Cell(Int32 row, Int32 column);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836035.aspx </remarks>
		/// <param name="beforeRow">object beforeRow</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Table Split(object beforeRow);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="nestedTables">optional object nestedTables</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ConvertToText(object separator, object nestedTables);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ConvertToText();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range ConvertToText(object separator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820953.aspx </remarks>
		/// <param name="behavior">NetOffice.WordApi.Enums.WdAutoFitBehavior behavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoFitBehavior(NetOffice.WordApi.Enums.WdAutoFitBehavior behavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192363.aspx </remarks>
		/// <param name="styleName">string styleName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyStyleDirectFormatting(string styleName);

		#endregion
	}
}
