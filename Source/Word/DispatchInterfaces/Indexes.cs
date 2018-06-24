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
	/// DispatchInterface Indexes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844981.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002097C-0000-0000-C000-000000000046")]
	public interface Indexes : ICOMObject, IEnumerableProvider<NetOffice.WordApi.Index>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834289.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822602.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195310.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839153.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840322.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdIndexFormat Format { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.Index this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		/// <param name="reading">optional object reading</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic, object reading);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841043.aspx </remarks>
		/// <param name="concordanceFileName">string concordanceFileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void AutoMarkEntries(string concordanceFileName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		/// <param name="indexLanguage">optional object indexLanguage</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy, object indexLanguage);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy);

        #endregion

        #region IEnumerable<NetOffice.WordApi.Index>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        new  IEnumerator<NetOffice.WordApi.Index> GetEnumerator();

        #endregion
    }
}
