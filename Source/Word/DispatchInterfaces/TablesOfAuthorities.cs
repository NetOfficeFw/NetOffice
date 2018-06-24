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
	/// DispatchInterface TablesOfAuthorities 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837712.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("00020912-0000-0000-C000-000000000046")]
	public interface TablesOfAuthorities : ICOMObject, IEnumerableProvider<NetOffice.WordApi.TableOfAuthorities>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845059.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838690.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837691.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdToaFormat Format { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.TableOfAuthorities this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object includeCategoryHeader</param>
		/// <param name="pageNumberSeparator">optional object pageNumberSeparator</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader, object pageNumberSeparator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object includeCategoryHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837703.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void NextCitation(string shortCitation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText, object category);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText, object category);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllCitations(string shortCitation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllCitations(string shortCitation, object longCitation);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText);

        #endregion

        #region IEnumerable<NetOffice.WordApi.TableOfAuthorities>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.TableOfAuthorities> GetEnumerator();

        #endregion
    }
}
