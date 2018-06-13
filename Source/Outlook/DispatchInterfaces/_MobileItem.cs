using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _MobileItem 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000630FE-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.MobileItem))]
    public interface _MobileItem : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Actions Actions { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Attachments Attachments { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string BillingInformation { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string Body { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string Categories { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string Companies { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationIndex { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationTopic { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		DateTime CreationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.FormDescription FormDescription { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Inspector GetInspector { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlImportance Importance { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		DateTime LastModificationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object MAPIOBJECT { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string MessageClass { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string Mileage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		bool NoAging { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		Int32 OutlookInternalVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string OutlookVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		bool Saved { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlSensitivity Sensitivity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		Int32 Size { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string Subject { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		bool UnRead { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.UserProperties UserProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string HTMLBody { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlMobileFormat MobileFormat { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string SMILBody { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Recipients Recipients { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string To { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string ReplyRecipientNames { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Recipients ReplyRecipients { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		bool Submitted { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.ItemProperties ItemProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		DateTime ReceivedTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Account SendUsingAccount { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		bool Sent { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		DateTime SentOn { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.PropertyAccessor PropertyAccessor { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string ReceivedByEntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string ReceivedByName { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string SenderEmailAddress { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string SenderEmailType { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		string SenderName { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		object Copy();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void Display(object modal);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 14,15,16)]
		void Display();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 14,15,16)]
		object Move(NetOffice.OutlookApi.MAPIFolder destFldr);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void SaveAs(string path, object type);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 14,15,16)]
		void SaveAs(string path);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.MobileItem Reply();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.MobileItem ReplyAll();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="forceSend">bool forceSend</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void Send(bool forceSend);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.MobileItem Forward();

		#endregion
	}
}
