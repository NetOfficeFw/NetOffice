using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface MAPIFolder 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class MAPIFolder : COMObject
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
                    _type = typeof(MAPIFolder);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MAPIFolder(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MAPIFolder(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MAPIFolder(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlItemType DefaultItemType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlItemType>(this, "DefaultItemType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string DefaultMessageClass
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultMessageClass");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Description
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Items Items
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Items>(this, "Items");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string StoreID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StoreID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 UnReadItemCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UnReadItemCount");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object UserPermissions
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "UserPermissions");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool WebViewOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebViewOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebViewOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string WebViewURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WebViewURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebViewURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool WebViewAllowNavigation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebViewAllowNavigation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebViewAllowNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public string AddressBookName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AddressBookName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddressBookName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool ShowAsOutlookAB
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAsOutlookAB");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAsOutlookAB", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public string FolderPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool InAppFolderSyncObject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InAppFolderSyncObject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InAppFolderSyncObject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.View CurrentView
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.View>(this, "CurrentView", NetOffice.OutlookApi.View.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool CustomViewsOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CustomViewsOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CustomViewsOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Views Views
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Views>(this, "Views");
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
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FullFolderPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullFolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public bool IsSharePointFolder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSharePointFolder");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlShowItemCount ShowItemCount
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlShowItemCount>(this, "ShowItemCount");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowItemCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Store Store
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Store>(this, "Store", NetOffice.OutlookApi.Store.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", NetOffice.OutlookApi.PropertyAccessor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.UserDefinedProperties UserDefinedProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserDefinedProperties>(this, "UserDefinedProperties", NetOffice.OutlookApi.UserDefinedProperties.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="destinationFolder">NetOffice.OutlookApi.MAPIFolder destinationFolder</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder CopyTo(NetOffice.OutlookApi.MAPIFolder destinationFolder)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "CopyTo", destinationFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="displayMode">optional object displayMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Explorer GetExplorer(object displayMode)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(this, "GetExplorer", displayMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Explorer GetExplorer()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(this, "GetExplorer");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="destinationFolder">NetOffice.OutlookApi.MAPIFolder destinationFolder</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void MoveTo(NetOffice.OutlookApi.MAPIFolder destinationFolder)
		{
			 Factory.ExecuteMethod(this, "MoveTo", destinationFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void AddToPFFavorites()
		{
			 Factory.ExecuteMethod(this, "AddToPFFavorites");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fNoUI">optional object fNoUI</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void AddToFavorites(object fNoUI, object name)
		{
			 Factory.ExecuteMethod(this, "AddToFavorites", fNoUI, name);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			 Factory.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fNoUI">optional object fNoUI</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void AddToFavorites(object fNoUI)
		{
			 Factory.ExecuteMethod(this, "AddToFavorites", fNoUI);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="storageIdentifier">string storageIdentifier</param>
		/// <param name="storageIdentifierType">NetOffice.OutlookApi.Enums.OlStorageIdentifierType storageIdentifierType</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._StorageItem GetStorage(string storageIdentifier, NetOffice.OutlookApi.Enums.OlStorageIdentifierType storageIdentifierType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._StorageItem>(this, "GetStorage", storageIdentifier, storageIdentifierType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="filter">optional object filter</param>
		/// <param name="tableContents">optional object tableContents</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Table GetTable(object filter, object tableContents)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", NetOffice.OutlookApi.Table.LateBindingApiWrapperType, filter, tableContents);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Table GetTable()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", NetOffice.OutlookApi.Table.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Table GetTable(object filter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", NetOffice.OutlookApi.Table.LateBindingApiWrapperType, filter);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.CalendarSharing GetCalendarExporter()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.CalendarSharing>(this, "GetCalendarExporter", NetOffice.OutlookApi.CalendarSharing.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), NativeResult]
		public stdole.Picture GetCustomIcon()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCustomIcon", paramsArray);
            return returnItem as stdole.Picture;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="picture">stdole.Picture picture</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void SetCustomIcon(stdole.Picture picture)
		{
			 Factory.ExecuteMethod(this, "SetCustomIcon", picture);
		}

		#endregion

		#pragma warning restore
	}
}
