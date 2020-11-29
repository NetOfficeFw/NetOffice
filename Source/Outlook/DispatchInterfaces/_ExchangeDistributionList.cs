using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _ExchangeDistributionList 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _ExchangeDistributionList : COMObject
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
                    _type = typeof(_ExchangeDistributionList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _ExchangeDistributionList(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ExchangeDistributionList(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeDistributionList(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Application"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Class"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Session"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Parent"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Address"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Address");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.DisplayType"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDisplayType DisplayType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDisplayType>(this, "DisplayType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.ID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OutlookApi.AddressEntry Manager
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntry>(this, "Manager", NetOffice.OutlookApi.AddressEntry.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MAPIOBJECT
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "MAPIOBJECT", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OutlookApi.AddressEntries Members
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntries>(this, "Members", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Name"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Type"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Type
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Type");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.AddressEntryUserType"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAddressEntryUserType AddressEntryUserType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlAddressEntryUserType>(this, "AddressEntryUserType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.PropertyAccessor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Alias"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Alias
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Alias");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Comments"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Comments
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Comments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Comments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.PrimarySmtpAddress"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string PrimarySmtpAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrimarySmtpAddress");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Delete"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Details"/> </remarks>
		/// <param name="hWnd">optional object hWnd</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Details(object hWnd)
		{
			 Factory.ExecuteMethod(this, "Details", hWnd);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Details"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Details()
		{
			 Factory.ExecuteMethod(this, "Details");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetFreeBusy"/> </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		/// <param name="completeFormat">optional object completeFormat</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string GetFreeBusy(DateTime start, Int32 minPerChar, object completeFormat)
		{
			return Factory.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar, completeFormat);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetFreeBusy"/> </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string GetFreeBusy(DateTime start, Int32 minPerChar)
		{
			return Factory.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Update"/> </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		/// <param name="refresh">optional object refresh</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Update(object makePermanent, object refresh)
		{
			 Factory.ExecuteMethod(this, "Update", makePermanent, refresh);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Update"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.Update"/> </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Update(object makePermanent)
		{
			 Factory.ExecuteMethod(this, "Update", makePermanent);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void UpdateFreeBusy()
		{
			 Factory.ExecuteMethod(this, "UpdateFreeBusy");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetContact"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._ContactItem GetContact()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._ContactItem>(this, "GetContact");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetExchangeUser"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ExchangeUser GetExchangeUser()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUser", NetOffice.OutlookApi.ExchangeUser.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetExchangeDistributionList"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ExchangeDistributionList GetExchangeDistributionList()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeDistributionList>(this, "GetExchangeDistributionList", NetOffice.OutlookApi.ExchangeDistributionList.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetMemberOfList"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntries GetMemberOfList()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetMemberOfList", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetExchangeDistributionListMembers"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntries GetExchangeDistributionListMembers()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetExchangeDistributionListMembers", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList.GetOwners"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntries GetOwners()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetOwners", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
