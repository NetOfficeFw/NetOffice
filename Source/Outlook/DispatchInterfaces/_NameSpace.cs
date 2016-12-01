using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _NameSpace 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _NameSpace : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NameSpace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869036.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OutlookApi._Application newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869498.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Class", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlObjectClass)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867831.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Session", paramsArray);
				NetOffice.OutlookApi._NameSpace newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869215.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869341.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient CurrentUser
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentUser", paramsArray);
				NetOffice.OutlookApi.Recipient newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType) as NetOffice.OutlookApi.Recipient;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868379.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Folders Folders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Folders", paramsArray);
				NetOffice.OutlookApi._Folders newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Folders;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868360.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public string Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866049.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.AddressLists AddressLists
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddressLists", paramsArray);
				NetOffice.OutlookApi.AddressLists newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.AddressLists.LateBindingApiWrapperType) as NetOffice.OutlookApi.AddressLists;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860705.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.SyncObjects SyncObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SyncObjects", paramsArray);
				NetOffice.OutlookApi.SyncObjects newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.SyncObjects.LateBindingApiWrapperType) as NetOffice.OutlookApi.SyncObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869001.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public bool Offline
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Offline", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MAPIOBJECT
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MAPIOBJECT", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864709.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlExchangeConnectionMode ExchangeConnectionMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExchangeConnectionMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlExchangeConnectionMode)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867142.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Accounts Accounts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Accounts", paramsArray);
				NetOffice.OutlookApi.Accounts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Accounts.LateBindingApiWrapperType) as NetOffice.OutlookApi.Accounts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866464.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public string CurrentProfileName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentProfileName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864790.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Stores Stores
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stores", paramsArray);
				NetOffice.OutlookApi.Stores newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Stores.LateBindingApiWrapperType) as NetOffice.OutlookApi.Stores;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863979.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Store DefaultStore
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultStore", paramsArray);
				NetOffice.OutlookApi.Store newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Store.LateBindingApiWrapperType) as NetOffice.OutlookApi.Store;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863625.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Categories Categories
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Categories", paramsArray);
				NetOffice.OutlookApi.Categories newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Categories.LateBindingApiWrapperType) as NetOffice.OutlookApi.Categories;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860362.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public string ExchangeMailboxServerName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExchangeMailboxServerName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860350.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public string ExchangeMailboxServerVersion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExchangeMailboxServerVersion", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863351.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public string AutoDiscoverXml
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoDiscoverXml", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868381.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAutoDiscoverConnectionMode AutoDiscoverConnectionMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoDiscoverConnectionMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlAutoDiscoverConnectionMode)intReturnItem;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866418.aspx
		/// </summary>
		/// <param name="recipientName">string RecipientName</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient CreateRecipient(string recipientName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipientName);
			object returnItem = Invoker.MethodReturn(this, "CreateRecipient", paramsArray);
			NetOffice.OutlookApi.Recipient newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType) as NetOffice.OutlookApi.Recipient;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866724.aspx
		/// </summary>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders FolderType</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetDefaultFolder(NetOffice.OutlookApi.Enums.OlDefaultFolders folderType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(folderType);
			object returnItem = Invoker.MethodReturn(this, "GetDefaultFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx
		/// </summary>
		/// <param name="entryIDFolder">string EntryIDFolder</param>
		/// <param name="entryIDStore">optional object EntryIDStore</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder, object entryIDStore)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(entryIDFolder, entryIDStore);
			object returnItem = Invoker.MethodReturn(this, "GetFolderFromID", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861238.aspx
		/// </summary>
		/// <param name="entryIDFolder">string EntryIDFolder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetFolderFromID(string entryIDFolder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(entryIDFolder);
			object returnItem = Invoker.MethodReturn(this, "GetFolderFromID", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx
		/// </summary>
		/// <param name="entryIDItem">string EntryIDItem</param>
		/// <param name="entryIDStore">optional object EntryIDStore</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object GetItemFromID(string entryIDItem, object entryIDStore)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(entryIDItem, entryIDStore);
			object returnItem = Invoker.MethodReturn(this, "GetItemFromID", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869893.aspx
		/// </summary>
		/// <param name="entryIDItem">string EntryIDItem</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object GetItemFromID(string entryIDItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(entryIDItem);
			object returnItem = Invoker.MethodReturn(this, "GetItemFromID", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867251.aspx
		/// </summary>
		/// <param name="entryID">string EntryID</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipient GetRecipientFromID(string entryID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(entryID);
			object returnItem = Invoker.MethodReturn(this, "GetRecipientFromID", paramsArray);
			NetOffice.OutlookApi.Recipient newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Recipient.LateBindingApiWrapperType) as NetOffice.OutlookApi.Recipient;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869575.aspx
		/// </summary>
		/// <param name="recipient">NetOffice.OutlookApi.Recipient Recipient</param>
		/// <param name="folderType">NetOffice.OutlookApi.Enums.OlDefaultFolders FolderType</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetSharedDefaultFolder(NetOffice.OutlookApi.Recipient recipient, NetOffice.OutlookApi.Enums.OlDefaultFolders folderType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipient, folderType);
			object returnItem = Invoker.MethodReturn(this, "GetSharedDefaultFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870076.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logoff()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Logoff", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx
		/// </summary>
		/// <param name="profile">optional object Profile</param>
		/// <param name="password">optional object Password</param>
		/// <param name="showDialog">optional object ShowDialog</param>
		/// <param name="newSession">optional object NewSession</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password, object showDialog, object newSession)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(profile, password, showDialog, newSession);
			Invoker.Method(this, "Logon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logon()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Logon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx
		/// </summary>
		/// <param name="profile">optional object Profile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(profile);
			Invoker.Method(this, "Logon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx
		/// </summary>
		/// <param name="profile">optional object Profile</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(profile, password);
			Invoker.Method(this, "Logon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861594.aspx
		/// </summary>
		/// <param name="profile">optional object Profile</param>
		/// <param name="password">optional object Password</param>
		/// <param name="showDialog">optional object ShowDialog</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Logon(object profile, object password, object showDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(profile, password, showDialog);
			Invoker.Method(this, "Logon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869969.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder PickFolder()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PickFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void RefreshRemoteHeaders()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshRemoteHeaders", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869071.aspx
		/// </summary>
		/// <param name="store">object Store</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void AddStore(object store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			Invoker.Method(this, "AddStore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864232.aspx
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder Folder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void RemoveStore(NetOffice.OutlookApi.MAPIFolder folder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(folder);
			Invoker.Method(this, "RemoveStore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx
		/// </summary>
		/// <param name="contactItem">optional object ContactItem</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public void Dial(object contactItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(contactItem);
			Invoker.Method(this, "Dial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862163.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public void Dial()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861574.aspx
		/// </summary>
		/// <param name="store">object Store</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlStoreType Type</param>
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		public void AddStoreEx(object store, NetOffice.OutlookApi.Enums.OlStoreType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store, type);
			Invoker.Method(this, "AddStoreEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867382.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SelectNamesDialog GetSelectNamesDialog()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetSelectNamesDialog", paramsArray);
			NetOffice.OutlookApi.SelectNamesDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.SelectNamesDialog.LateBindingApiWrapperType) as NetOffice.OutlookApi.SelectNamesDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861834.aspx
		/// </summary>
		/// <param name="showProgressDialog">bool showProgressDialog</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public void SendAndReceive(bool showProgressDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showProgressDialog);
			Invoker.Method(this, "SendAndReceive", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860428.aspx
		/// </summary>
		/// <param name="iD">string ID</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntry GetAddressEntryFromID(string iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD);
			object returnItem = Invoker.MethodReturn(this, "GetAddressEntryFromID", paramsArray);
			NetOffice.OutlookApi.AddressEntry newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.AddressEntry.LateBindingApiWrapperType) as NetOffice.OutlookApi.AddressEntry;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860951.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressList GetGlobalAddressList()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetGlobalAddressList", paramsArray);
			NetOffice.OutlookApi.AddressList newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.AddressList.LateBindingApiWrapperType) as NetOffice.OutlookApi.AddressList;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868771.aspx
		/// </summary>
		/// <param name="iD">string ID</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Store GetStoreFromID(string iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD);
			object returnItem = Invoker.MethodReturn(this, "GetStoreFromID", paramsArray);
			NetOffice.OutlookApi.Store newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Store.LateBindingApiWrapperType) as NetOffice.OutlookApi.Store;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="name">optional object Name</param>
		/// <param name="downloadAttachments">optional object DownloadAttachments</param>
		/// <param name="useTTL">optional object UseTTL</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments, object useTTL)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, downloadAttachments, useTTL);
			object returnItem = Invoker.MethodReturn(this, "OpenSharedFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			object returnItem = Invoker.MethodReturn(this, "OpenSharedFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name);
			object returnItem = Invoker.MethodReturn(this, "OpenSharedFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867648.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="name">optional object Name</param>
		/// <param name="downloadAttachments">optional object DownloadAttachments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder(string path, object name, object downloadAttachments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, downloadAttachments);
			object returnItem = Invoker.MethodReturn(this, "OpenSharedFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869733.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object OpenSharedItem(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			object returnItem = Invoker.MethodReturn(this, "OpenSharedItem", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx
		/// </summary>
		/// <param name="context">object Context</param>
		/// <param name="provider">optional object Provider</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem CreateSharingItem(object context, object provider)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(context, provider);
			object returnItem = Invoker.MethodReturn(this, "CreateSharingItem", paramsArray);
			NetOffice.OutlookApi.SharingItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType) as NetOffice.OutlookApi.SharingItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864725.aspx
		/// </summary>
		/// <param name="context">object Context</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem CreateSharingItem(object context)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(context);
			object returnItem = Invoker.MethodReturn(this, "CreateSharingItem", paramsArray);
			NetOffice.OutlookApi.SharingItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType) as NetOffice.OutlookApi.SharingItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864762.aspx
		/// </summary>
		/// <param name="firstEntryID">string FirstEntryID</param>
		/// <param name="secondEntryID">string SecondEntryID</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public bool CompareEntryIDs(string firstEntryID, string secondEntryID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(firstEntryID, secondEntryID);
			object returnItem = Invoker.MethodReturn(this, "CompareEntryIDs", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869218.aspx
		/// </summary>
		/// <param name="addressEntry">NetOffice.OutlookApi.AddressEntry AddressEntry</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OfficeApi.ContactCard CreateContactCard(NetOffice.OutlookApi.AddressEntry addressEntry)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(addressEntry);
			object returnItem = Invoker.MethodReturn(this, "CreateContactCard", paramsArray);
			NetOffice.OfficeApi.ContactCard newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.ContactCard.LateBindingApiWrapperType) as NetOffice.OfficeApi.ContactCard;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}