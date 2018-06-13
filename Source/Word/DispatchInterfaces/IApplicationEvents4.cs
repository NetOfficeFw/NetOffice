using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents4 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00020A01-0001-0000-C000-000000000046")]
	public interface IApplicationEvents4 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Startup();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentChange();

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentOpen(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void NewDocument(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowSelectionChange(NetOffice.WordApi.Selection sel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void EPostagePropertyDialog(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void EPostageInsert(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="docResult">NetOffice.WordApi.Document docResult</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeBeforeMerge(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeDataSourceLoad(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, bool handled);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		/// <param name="toState">Int32 toState</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void MailMergeWizardStateChange(NetOffice.WordApi.Document doc, Int32 fromState, Int32 toState, bool handled);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="oldXMLNode">NetOffice.WordApi.XMLNode oldXMLNode</param>
		/// <param name="newXMLNode">NetOffice.WordApi.XMLNode newXMLNode</param>
		/// <param name="reason">Int32 reason</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void XMLSelectionChange(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, Int32 reason);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLNode">NetOffice.WordApi.XMLNode xMLNode</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void XMLValidationError(NetOffice.WordApi.XMLNode xMLNode);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void DocumentSync(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cpDeliveryAddrStart">Int32 cpDeliveryAddrStart</param>
		/// <param name="cpDeliveryAddrEnd">Int32 cpDeliveryAddrEnd</param>
		/// <param name="cpReturnAddrStart">Int32 cpReturnAddrStart</param>
		/// <param name="cpReturnAddrEnd">Int32 cpReturnAddrEnd</param>
		/// <param name="xaWidth">Int32 xaWidth</param>
		/// <param name="yaHeight">Int32 yaHeight</param>
		/// <param name="bstrPrinterName">string bstrPrinterName</param>
		/// <param name="bstrPaperFeed">string bstrPaperFeed</param>
		/// <param name="fPrint">bool fPrint</param>
		/// <param name="fCancel">bool fCancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		void EPostageInsertEx(NetOffice.WordApi.Document doc, Int32 cpDeliveryAddrStart, Int32 cpDeliveryAddrEnd, Int32 cpReturnAddrStart, Int32 cpReturnAddrEnd, Int32 xaWidth, Int32 yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, bool fCancel);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 MailMergeDataSourceValidate2(NetOffice.WordApi.Document doc, bool handled);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowOpen(NetOffice.WordApi.ProtectedViewWindow pvWindow);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowBeforeEdit(NetOffice.WordApi.ProtectedViewWindow pvWindow, bool cancel);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		/// <param name="closeReason">Int32 closeReason</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowBeforeClose(NetOffice.WordApi.ProtectedViewWindow pvWindow, Int32 closeReason, bool cancel);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowSize(NetOffice.WordApi.ProtectedViewWindow pvWindow);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowActivate(NetOffice.WordApi.ProtectedViewWindow pvWindow);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		Int32 ProtectedViewWindowDeactivate(NetOffice.WordApi.ProtectedViewWindow pvWindow);

		#endregion
	}
}
