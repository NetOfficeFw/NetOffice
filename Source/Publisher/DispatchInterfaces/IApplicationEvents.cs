using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002123F-0000-0000-C000-000000000046")]
	public interface IApplicationEvents : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window wn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void WindowActivate(NetOffice.PublisherApi.Window wn);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window wn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void WindowDeactivate(NetOffice.PublisherApi.Window wn);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="vw">NetOffice.PublisherApi.View vw</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void WindowPageChange(NetOffice.PublisherApi.View vw);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void NewDocument(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void DocumentOpen(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void DocumentBeforeClose(NetOffice.PublisherApi._Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeAfterMerge(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeAfterRecordMerge(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeBeforeMerge(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, bool cancel);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeBeforeRecordMerge(NetOffice.PublisherApi._Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeDataSourceLoad(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeWizardSendToCustom(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeWizardStateChange(NetOffice.PublisherApi._Document doc, Int32 fromState);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeDataSourceValidate(NetOffice.PublisherApi._Document doc, bool handled);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="okToInsert">bool okToInsert</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeInsertBarcode(NetOffice.PublisherApi._Document doc, bool okToInsert);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeRecipientListClose(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="bstrString">string bstrString</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeGenerateBarcode(NetOffice.PublisherApi._Document doc, string bstrString);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MailMergeWizardFollowUpCustom(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void BeforePrint(NetOffice.PublisherApi._Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void AfterPrint(NetOffice.PublisherApi._Document doc);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowCatalogUI();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void HideCatalogUI();

		#endregion
	}
}
