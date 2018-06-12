using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ListFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820923.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ListFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844779.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ListLevelNumber { get; set; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839514.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.List List { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821156.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.ListTemplate ListTemplate { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196311.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 ListValue { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836742.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SingleList { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835216.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		bool SingleListTemplate { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197765.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdListType ListType { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836909.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		string ListString { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194491.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195720.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194230.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837259.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.InlineShape ListPictureBullet { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196639.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		NetOffice.WordApi.Enums.WdContinue CanContinuePreviousList(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821874.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void RemoveNumbers();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196580.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ConvertNumbersToText();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType, object level);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820721.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		Int32 CountNumberedItems(object numberType);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyBulletDefaultOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyNumberDefaultOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyOutlineNumberDefaultOld();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo);

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
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplateOld(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840622.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ListOutdent();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192794.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ListIndent();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyBulletDefault(object defaultListBehavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194355.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyBulletDefault();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyNumberDefault(object defaultListBehavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839320.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyNumberDefault();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx </remarks>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyOutlineNumberDefault(object defaultListBehavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822954.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyOutlineNumberDefault();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197184.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void ApplyListTemplate(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		/// <param name="applyLevel">optional object applyLevel</param>
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior, object applyLevel);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195913.aspx </remarks>
		/// <param name="listTemplate">NetOffice.WordApi.ListTemplate listTemplate</param>
		/// <param name="continuePreviousList">optional object continuePreviousList</param>
		/// <param name="applyTo">optional object applyTo</param>
		/// <param name="defaultListBehavior">optional object defaultListBehavior</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		void ApplyListTemplateWithLevel(NetOffice.WordApi.ListTemplate listTemplate, object continuePreviousList, object applyTo, object defaultListBehavior);

		#endregion
	}
}
