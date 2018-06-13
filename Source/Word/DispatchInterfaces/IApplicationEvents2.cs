using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents2 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000209FE-0001-0000-C000-000000000046")]
	public interface IApplicationEvents2 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Startup();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DocumentChange();

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DocumentOpen(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void NewDocument(NetOffice.WordApi.Document doc);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WindowSelectionChange(NetOffice.WordApi.Selection sel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel);

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel);

		#endregion
	}
}
