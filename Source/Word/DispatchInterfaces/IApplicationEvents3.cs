using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents3 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface IApplicationEvents3 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Startup();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DocumentChange();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DocumentOpen(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void NewDocument(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowSelectionChange(NetOffice.WordApi.Selection sel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void EPostagePropertyDialog(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void EPostageInsert(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="docResult">NetOffice.WordApi.Document docResult</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeBeforeMerge(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeDataSourceLoad(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, bool handled);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		/// <param name="toState">Int32 toState</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MailMergeWizardStateChange(NetOffice.WordApi.Document doc, Int32 fromState, Int32 toState, bool handled);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		#endregion
	}
}
