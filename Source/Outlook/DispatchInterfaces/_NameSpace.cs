using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _NameSpace 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063002-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.NameSpace))]
    public interface _NameSpace : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869036.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869498.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867831.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869215.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869341.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Recipient CurrentUser { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868379.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Folders Folders { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868360.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Type { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866049.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.AddressLists AddressLists { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860705.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.SyncObjects SyncObjects { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869001.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		bool Offline { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object MAPIOBJECT { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864709.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlExchangeConnectionMode ExchangeConnectionMode { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867142.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Accounts Accounts { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866464.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string CurrentProfileName { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864790.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Stores Stores { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863979.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Store DefaultStore { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863625.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Categories Categories { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860362.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string ExchangeMailboxServerName { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860350.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string ExchangeMailboxServerVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863351.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string AutoDiscoverXml { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868381.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlAutoDiscoverConnectionMode AutoDiscoverConnectionMode { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866418.aspx </remarks>
		/// <param name="recipientName">string recipientName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Recipient CreateRecipient(string recipientName);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866724.aspx </remarks>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders folderType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder GetDefaultFolder(NetOffice.OutlookApi.Enums.OlDefaultFolders folderType);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx </remarks>
		/// <param name="entryIDFolder">string entryIDFolder</param>
		/// <param name="entryIDStore">optional object entryIDStore</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder, object entryIDStore);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx </remarks>
		/// <param name="entryIDFolder">string entryIDFolder</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx </remarks>
		/// <param name="entryIDItem">string entryIDItem</param>
		/// <param name="entryIDStore">optional object entryIDStore</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object GetItemFromID(string entryIDItem, object entryIDStore);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx </remarks>
		/// <param name="entryIDItem">string entryIDItem</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object GetItemFromID(string entryIDItem);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867251.aspx </remarks>
		/// <param name="entryID">string entryID</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Recipient GetRecipientFromID(string entryID);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869575.aspx </remarks>
		/// <param name="recipient">NetOffice.OutlookApi.Recipient recipient</param>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders folderType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder GetSharedDefaultFolder(NetOffice.OutlookApi.Recipient recipient, NetOffice.OutlookApi.Enums.OlDefaultFolders folderType);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870076.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logoff();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		/// <param name="showDialog">optional object showDialog</param>
		/// <param name="newSession">optional object newSession</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logon(object profile, object password, object showDialog, object newSession);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logon();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logon(object profile);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logon(object profile, object password);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		/// <param name="showDialog">optional object showDialog</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Logon(object profile, object password, object showDialog);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869969.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder PickFolder();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void RefreshRemoteHeaders();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869071.aspx </remarks>
		/// <param name="store">object store</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void AddStore(object store);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864232.aspx </remarks>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void RemoveStore(NetOffice.OutlookApi.MAPIFolder folder);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx </remarks>
		/// <param name="contactItem">optional object contactItem</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void Dial(object contactItem);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void Dial();

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861574.aspx </remarks>
		/// <param name="store">object store</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlStoreType type</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		void AddStoreEx(object store, NetOffice.OutlookApi.Enums.OlStoreType type);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867382.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.SelectNamesDialog GetSelectNamesDialog();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861834.aspx </remarks>
		/// <param name="showProgressDialog">bool showProgressDialog</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void SendAndReceive(bool showProgressDialog);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860428.aspx </remarks>
		/// <param name="iD">string iD</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.AddressEntry GetAddressEntryFromID(string iD);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860951.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.AddressList GetGlobalAddressList();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868771.aspx </remarks>
		/// <param name="iD">string iD</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Store GetStoreFromID(string iD);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="name">optional object name</param>
		/// <param name="downloadAttachments">optional object downloadAttachments</param>
		/// <param name="useTTL">optional object useTTL</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments, object useTTL);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="name">optional object name</param>
		/// <param name="downloadAttachments">optional object downloadAttachments</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869733.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object OpenSharedItem(string path);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx </remarks>
		/// <param name="context">object context</param>
		/// <param name="provider">optional object provider</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.SharingItem CreateSharingItem(object context, object provider);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx </remarks>
		/// <param name="context">object context</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.SharingItem CreateSharingItem(object context);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864762.aspx </remarks>
		/// <param name="firstEntryID">string firstEntryID</param>
		/// <param name="secondEntryID">string secondEntryID</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool CompareEntryIDs(string firstEntryID, string secondEntryID);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869218.aspx </remarks>
		/// <param name="addressEntry">NetOffice.OutlookApi.AddressEntry addressEntry</param>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OfficeApi.ContactCard CreateContactCard(NetOffice.OutlookApi.AddressEntry addressEntry);

		#endregion
	}
}
