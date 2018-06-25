using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Pages 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00021248-0000-0000-C000-000000000046")]
	public interface Pages : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Page>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="item">Int32 item</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Page this[Int32 item] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after, object duplicateObjectsOnPage);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddWizardPage10(Int32 after, object pageType);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddWizardPage10(Int32 after);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pageID">Int32 pageID</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page FindByPageID(Int32 pageID);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage, object addHyperlinkToWebNavBar);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page Add(Int32 count, Int32 after);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void AddWizardPage(Int32 after, object pageType, object addHyperlinkToWebNavBar);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddWizardPage(Int32 after);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void AddWizardPage(Int32 after, object pageType);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Page>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Page> GetEnumerator();

        #endregion
    }
}
