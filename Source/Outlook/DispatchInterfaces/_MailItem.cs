using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _MailItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063034-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.MailItem))]
    public interface _MailItem : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869350.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866957.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864227.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863655.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861914.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Actions Actions { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866435.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Attachments Attachments { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869243.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BillingInformation { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865304.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Body { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860423.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Categories { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861903.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Companies { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869408.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ConversationIndex { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869318.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ConversationTopic { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867230.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime CreationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866458.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860627.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.FormDescription FormDescription { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868098.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Inspector GetInspector { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866759.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlImportance Importance { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867677.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime LastModificationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object MAPIOBJECT { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867813.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MessageClass { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860348.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Mileage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869383.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool NoAging { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869069.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 OutlookInternalVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868956.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OutlookVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865073.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool Saved { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868972.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlSensitivity Sensitivity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861257.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Size { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865652.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Subject { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868556.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool UnRead { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866403.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.UserProperties UserProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868211.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool AlternateRecipientAllowed { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867162.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool AutoForwarded { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865864.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BCC { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869030.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CC { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869452.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime DeferredDeliveryTime { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868585.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool DeleteAfterSubmit { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861811.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime ExpiryTime { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime FlagDueBy { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861323.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FlagRequest { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlFlagStatus FlagStatus { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868941.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HTMLBody { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867402.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool OriginatorDeliveryReportRequested { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865400.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool ReadReceiptRequested { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869438.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReceivedByEntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866935.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReceivedByName { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870197.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReceivedOnBehalfOfEntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866908.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReceivedOnBehalfOfName { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867228.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime ReceivedTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870035.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool RecipientReassignmentProhibited { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865320.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Recipients Recipients { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866775.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool ReminderOverrideDefault { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867123.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool ReminderPlaySound { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870073.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool ReminderSet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861284.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReminderSoundFile { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868512.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime ReminderTime { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870011.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlRemoteStatus RemoteStatus { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867886.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReplyRecipientNames { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862985.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Recipients ReplyRecipients { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868473.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder SaveSentMessageFolder { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869598.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string SenderName { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868242.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool Sent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864408.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime SentOn { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862145.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string SentOnBehalfOfName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865326.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool Submitted { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860378.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string To { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866063.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string VotingOptions { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868303.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string VotingResponse { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Links Links { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865811.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.ItemProperties ItemProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869979.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlBodyFormat BodyFormat { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866978.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlDownloadState DownloadState { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860730.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		Int32 InternetCodepage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866920.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlRemoteStatus MarkForDownload { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865867.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		bool IsConflict { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool IsIPFax { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlFlagIcon FlagIcon { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool HasCoverSheet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863715.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		bool AutoResolvedWinner { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862967.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Conflicts Conflicts { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868262.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		string SenderEmailAddress { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869674.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		string SenderEmailType { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool EnableSharedAttachments { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863622.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlPermission Permission { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869080.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlPermissionService PermissionService { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868823.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.PropertyAccessor PropertyAccessor { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869311.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Account SendUsingAccount { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870037.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string TaskSubject { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861586.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskDueDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866742.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskStartDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864714.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskCompletedDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869249.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime ToDoTaskOrdinal { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866239.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool IsMarkedAsTask { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867895.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationID { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869056.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.AddressEntry Sender { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863315.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string PermissionTemplateGuid { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867828.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		object RTFBody { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862673.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string RetentionPolicyName { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867620.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		DateTime RetentionExpirationDate { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860308.aspx </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868420.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object Copy();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863343.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861853.aspx </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display(object modal);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861853.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860683.aspx </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object Move(NetOffice.OutlookApi.MAPIFolder destFldr);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861582.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866979.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868727.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void SaveAs(string path, object type);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868727.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void SaveAs(string path);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865035.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void ClearConversationIndex();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865399.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.MailItem Forward();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868875.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.MailItem Reply();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862498.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.MailItem ReplyAll();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866779.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Send();

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862218.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void ShowCategoriesDialog();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868298.aspx </remarks>
		/// <param name="contact">NetOffice.OutlookApi.ContactItem contact</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void AddBusinessCard(NetOffice.OutlookApi.ContactItem contact);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869791.aspx </remarks>
		/// <param name="markInterval">NetOffice.OutlookApi.Enums.OlMarkInterval markInterval</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void MarkAsTask(NetOffice.OutlookApi.Enums.OlMarkInterval markInterval);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867188.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void ClearTaskFlag();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869870.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Conversation GetConversation();

		#endregion
	}
}
