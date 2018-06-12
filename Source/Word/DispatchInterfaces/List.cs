using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface List 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192789.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface List : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840272.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Range Range { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193376.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ListParagraphs ListParagraphs { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838053.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SingleListTemplate { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194676.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836145.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194035.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194057.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		string StyleName { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193093.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838078.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType, object level);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820795.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196826.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdContinue CanContinuePreviousList(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196090.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		/// <param name="applyLevel">optional object applyLevel</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior, object applyLevel);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191850.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object defaultListBehavior);

		#endregion
	}
}
