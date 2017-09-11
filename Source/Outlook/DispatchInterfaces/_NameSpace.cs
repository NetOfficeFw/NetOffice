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
 	public class _NameSpace : COMObject
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_NameSpace);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _NameSpace(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _NameSpace(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869036.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869498.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867831.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869215.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869341.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient CurrentUser
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipient>(this, "CurrentUser", NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868379.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Folders Folders
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Folders>(this, "Folders");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868360.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Type
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866049.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.AddressLists AddressLists
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressLists>(this, "AddressLists", NetOffice.OutlookApi.AddressLists.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860705.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.SyncObjects SyncObjects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.SyncObjects>(this, "SyncObjects", NetOffice.OutlookApi.SyncObjects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869001.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool Offline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Offline");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MAPIOBJECT
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864709.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlExchangeConnectionMode ExchangeConnectionMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlExchangeConnectionMode>(this, "ExchangeConnectionMode");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867142.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Accounts Accounts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Accounts>(this, "Accounts", NetOffice.OutlookApi.Accounts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866464.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string CurrentProfileName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentProfileName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864790.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Stores Stores
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Stores>(this, "Stores", NetOffice.OutlookApi.Stores.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863979.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Store DefaultStore
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Store>(this, "DefaultStore", NetOffice.OutlookApi.Store.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863625.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Categories Categories
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Categories>(this, "Categories", NetOffice.OutlookApi.Categories.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860362.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ExchangeMailboxServerName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ExchangeMailboxServerName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860350.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ExchangeMailboxServerVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ExchangeMailboxServerVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863351.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string AutoDiscoverXml
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AutoDiscoverXml");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868381.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAutoDiscoverConnectionMode AutoDiscoverConnectionMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlAutoDiscoverConnectionMode>(this, "AutoDiscoverConnectionMode");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866418.aspx </remarks>
		/// <param name="recipientName">string recipientName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient CreateRecipient(string recipientName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Recipient>(this, "CreateRecipient", NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType, recipientName);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866724.aspx </remarks>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders folderType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder GetDefaultFolder(NetOffice.OutlookApi.Enums.OlDefaultFolders folderType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetDefaultFolder", folderType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx </remarks>
		/// <param name="entryIDFolder">string entryIDFolder</param>
		/// <param name="entryIDStore">optional object entryIDStore</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder, object entryIDStore)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetFolderFromID", entryIDFolder, entryIDStore);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx </remarks>
		/// <param name="entryIDFolder">string entryIDFolder</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetFolderFromID", entryIDFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx </remarks>
		/// <param name="entryIDItem">string entryIDItem</param>
		/// <param name="entryIDStore">optional object entryIDStore</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object GetItemFromID(string entryIDItem, object entryIDStore)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetItemFromID", entryIDItem, entryIDStore);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx </remarks>
		/// <param name="entryIDItem">string entryIDItem</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object GetItemFromID(string entryIDItem)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetItemFromID", entryIDItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867251.aspx </remarks>
		/// <param name="entryID">string entryID</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient GetRecipientFromID(string entryID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Recipient>(this, "GetRecipientFromID", NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType, entryID);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869575.aspx </remarks>
		/// <param name="recipient">NetOffice.OutlookApi.Recipient recipient</param>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders folderType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder GetSharedDefaultFolder(NetOffice.OutlookApi.Recipient recipient, NetOffice.OutlookApi.Enums.OlDefaultFolders folderType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetSharedDefaultFolder", recipient, folderType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870076.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logoff()
		{
			 Factory.ExecuteMethod(this, "Logoff");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		/// <param name="showDialog">optional object showDialog</param>
		/// <param name="newSession">optional object newSession</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password, object showDialog, object newSession)
		{
			 Factory.ExecuteMethod(this, "Logon", profile, password, showDialog, newSession);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logon()
		{
			 Factory.ExecuteMethod(this, "Logon");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile)
		{
			 Factory.ExecuteMethod(this, "Logon", profile);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password)
		{
			 Factory.ExecuteMethod(this, "Logon", profile, password);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx </remarks>
		/// <param name="profile">optional object profile</param>
		/// <param name="password">optional object password</param>
		/// <param name="showDialog">optional object showDialog</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password, object showDialog)
		{
			 Factory.ExecuteMethod(this, "Logon", profile, password, showDialog);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869969.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder PickFolder()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "PickFolder");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void RefreshRemoteHeaders()
		{
			 Factory.ExecuteMethod(this, "RefreshRemoteHeaders");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869071.aspx </remarks>
		/// <param name="store">object store</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void AddStore(object store)
		{
			 Factory.ExecuteMethod(this, "AddStore", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864232.aspx </remarks>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void RemoveStore(NetOffice.OutlookApi.MAPIFolder folder)
		{
			 Factory.ExecuteMethod(this, "RemoveStore", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx </remarks>
		/// <param name="contactItem">optional object contactItem</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void Dial(object contactItem)
		{
			 Factory.ExecuteMethod(this, "Dial", contactItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void Dial()
		{
			 Factory.ExecuteMethod(this, "Dial");
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861574.aspx </remarks>
		/// <param name="store">object store</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlStoreType type</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void AddStoreEx(object store, NetOffice.OutlookApi.Enums.OlStoreType type)
		{
			 Factory.ExecuteMethod(this, "AddStoreEx", store, type);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867382.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SelectNamesDialog GetSelectNamesDialog()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SelectNamesDialog>(this, "GetSelectNamesDialog", NetOffice.OutlookApi.SelectNamesDialog.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861834.aspx </remarks>
		/// <param name="showProgressDialog">bool showProgressDialog</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SendAndReceive(bool showProgressDialog)
		{
			 Factory.ExecuteMethod(this, "SendAndReceive", showProgressDialog);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860428.aspx </remarks>
		/// <param name="iD">string iD</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntry GetAddressEntryFromID(string iD)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntry>(this, "GetAddressEntryFromID", NetOffice.OutlookApi.AddressEntry.LateBindingApiWrapperType, iD);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860951.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressList GetGlobalAddressList()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressList>(this, "GetGlobalAddressList", NetOffice.OutlookApi.AddressList.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868771.aspx </remarks>
		/// <param name="iD">string iD</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Store GetStoreFromID(string iD)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Store>(this, "GetStoreFromID", NetOffice.OutlookApi.Store.LateBindingApiWrapperType, iD);
		}

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
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments, object useTTL)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "OpenSharedFolder", path, name, downloadAttachments, useTTL);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "OpenSharedFolder", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "OpenSharedFolder", path, name);
		}

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
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "OpenSharedFolder", path, name, downloadAttachments);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869733.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object OpenSharedItem(string path)
		{
			return Factory.ExecuteVariantMethodGet(this, "OpenSharedItem", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx </remarks>
		/// <param name="context">object context</param>
		/// <param name="provider">optional object provider</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem CreateSharingItem(object context, object provider)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SharingItem>(this, "CreateSharingItem", NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType, context, provider);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx </remarks>
		/// <param name="context">object context</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem CreateSharingItem(object context)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SharingItem>(this, "CreateSharingItem", NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType, context);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864762.aspx </remarks>
		/// <param name="firstEntryID">string firstEntryID</param>
		/// <param name="secondEntryID">string secondEntryID</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool CompareEntryIDs(string firstEntryID, string secondEntryID)
		{
			return Factory.ExecuteBoolMethodGet(this, "CompareEntryIDs", firstEntryID, secondEntryID);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869218.aspx </remarks>
		/// <param name="addressEntry">NetOffice.OutlookApi.AddressEntry addressEntry</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OfficeApi.ContactCard CreateContactCard(NetOffice.OutlookApi.AddressEntry addressEntry)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.ContactCard>(this, "CreateContactCard", NetOffice.OfficeApi.ContactCard.LateBindingApiWrapperType, addressEntry);
		}

		#endregion

		#pragma warning restore
	}
}
