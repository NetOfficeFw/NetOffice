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
	/// DispatchInterface MailMergeFields 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835151.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002091F-0000-0000-C000-000000000046")]
	public interface MailMergeFields : ICOMObject, IEnumerableProvider<NetOffice.WordApi.MailMergeField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195330.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834563.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840148.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837178.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.MailMergeField this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836028.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField Add(NetOffice.WordApi.Range range, string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText, object askOnce);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText, object askOnce);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		/// <param name="falseText">optional object falseText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText, object falseText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195621.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddMergeRec(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839562.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddMergeSeq(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837747.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddNext(NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		/// <param name="valueAutoText">optional object valueAutoText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText, object valueAutoText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison);

        #endregion

        #region IEnumerable<NetOffice.WordApi.MailMergeField>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.MailMergeField> GetEnumerator();

        #endregion
    }
}
