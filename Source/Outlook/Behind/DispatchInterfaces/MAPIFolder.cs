using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface MAPIFolder 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class MAPIFolder : COMObject, NetOffice.OutlookApi.MAPIFolder
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OutlookApi.MAPIFolder);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MAPIFolder() : base()
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
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlItemType DefaultItemType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlItemType>(this, "DefaultItemType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string DefaultMessageClass
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultMessageClass");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string Description
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string EntryID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Folders Folders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Folders>(this, "Folders");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Items Items
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Items>(this, "Items");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string StoreID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StoreID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual Int32 UnReadItemCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UnReadItemCount");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object UserPermissions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "UserPermissions");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual bool WebViewOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WebViewOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WebViewOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string WebViewURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WebViewURL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WebViewURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual bool WebViewAllowNavigation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WebViewAllowNavigation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WebViewAllowNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual string AddressBookName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressBookName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddressBookName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual bool ShowAsOutlookAB
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowAsOutlookAB");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowAsOutlookAB", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual string FolderPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual bool InAppFolderSyncObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InAppFolderSyncObject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InAppFolderSyncObject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.View CurrentView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.View>(this, "CurrentView", typeof(NetOffice.OutlookApi.View));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual bool CustomViewsOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CustomViewsOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomViewsOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Views Views
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Views>(this, "Views");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object MAPIOBJECT
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string FullFolderPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullFolderPath");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual bool IsSharePointFolder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsSharePointFolder");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlShowItemCount ShowItemCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlShowItemCount>(this, "ShowItemCount");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ShowItemCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Store Store
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Store>(this, "Store", typeof(NetOffice.OutlookApi.Store));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", typeof(NetOffice.OutlookApi.PropertyAccessor));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.UserDefinedProperties UserDefinedProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserDefinedProperties>(this, "UserDefinedProperties", typeof(NetOffice.OutlookApi.UserDefinedProperties));
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
		public virtual NetOffice.OutlookApi.MAPIFolder CopyTo(NetOffice.OutlookApi.MAPIFolder destinationFolder)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "CopyTo", destinationFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Display()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="displayMode">optional object displayMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Explorer GetExplorer(object displayMode)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(this, "GetExplorer", displayMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi._Explorer GetExplorer()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(this, "GetExplorer");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="destinationFolder">NetOffice.OutlookApi.MAPIFolder destinationFolder</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void MoveTo(NetOffice.OutlookApi.MAPIFolder destinationFolder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveTo", destinationFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void AddToPFFavorites()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToPFFavorites");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fNoUI">optional object fNoUI</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual void AddToFavorites(object fNoUI, object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites", fNoUI, name);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual void AddToFavorites()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fNoUI">optional object fNoUI</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual void AddToFavorites(object fNoUI)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites", fNoUI);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="storageIdentifier">string storageIdentifier</param>
		/// <param name="storageIdentifierType">NetOffice.OutlookApi.Enums.OlStorageIdentifierType storageIdentifierType</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._StorageItem GetStorage(string storageIdentifier, NetOffice.OutlookApi.Enums.OlStorageIdentifierType storageIdentifierType)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._StorageItem>(this, "GetStorage", storageIdentifier, storageIdentifierType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="filter">optional object filter</param>
		/// <param name="tableContents">optional object tableContents</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Table GetTable(object filter, object tableContents)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", typeof(NetOffice.OutlookApi.Table), filter, tableContents);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Table GetTable()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", typeof(NetOffice.OutlookApi.Table));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Table GetTable(object filter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", typeof(NetOffice.OutlookApi.Table), filter);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.CalendarSharing GetCalendarExporter()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.CalendarSharing>(this, "GetCalendarExporter", typeof(NetOffice.OutlookApi.CalendarSharing));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), NativeResult]
		public virtual stdole.Picture GetCustomIcon()
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
		public virtual void SetCustomIcon(stdole.Picture picture)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetCustomIcon", picture);
		}

		#endregion

		#pragma warning restore
	}
}


