using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Workbooks 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841074.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000208DB-0000-0000-C000-000000000046")]
	public interface Workbooks : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.Workbook>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195019.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195436.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837124.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822893.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.Workbook this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx </remarks>
		/// <param name="template">optional object template</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Add(object template);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Add();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839657.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="local">optional object local</param>
		/// <param name="corruptLoad">optional object corruptLoad</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local, object corruptLoad);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="local">optional object local</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
		/// <param name="local">optional object local</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers, object local);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="importDataAs">optional object importDataAs</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery, object importDataAs);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenDatabase(string filename);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194062.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void CheckOut(string filename);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193284.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool CanCheckOut(string filename);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		/// <param name="loadOption">optional object loadOption</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets, object loadOption);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook OpenXML(string filename);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _OpenXML(string filename, object stylesheets);

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		NetOffice.ExcelApi.Workbook _OpenXML(string filename);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Workbook>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.Workbook> GetEnumerator();

        #endregion
    }
}
