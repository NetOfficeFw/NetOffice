using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TablesOfContents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838538.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("00020914-0000-0000-C000-000000000046")]
	public interface TablesOfContents : ICOMObject, IEnumerableProvider<NetOffice.WordApi.TableOfContents>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197427.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197796.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840817.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845238.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839904.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdTocFormat Format { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.TableOfContents this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID, object level);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		/// <param name="useOutlineLevels">optional object useOutlineLevels</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb, object useOutlineLevels);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks);

        #endregion

        #region IEnumerable<NetOffice.WordApi.TableOfContents>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.TableOfContents> GetEnumerator();

        #endregion
    }
}
