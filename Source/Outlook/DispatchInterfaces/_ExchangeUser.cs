using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _ExchangeUser 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _ExchangeUser : COMObject
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
                    _type = typeof(_ExchangeUser);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _ExchangeUser(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ExchangeUser(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ExchangeUser(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Class"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Session"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Address"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.DisplayType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.ID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Name"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Type"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.AddressEntryUserType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.PropertyAccessor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Alias"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.AssistantName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string AssistantName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AssistantName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AssistantName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.BusinessTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string BusinessTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.City"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string City
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "City");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "City", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Comments"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.CompanyName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string CompanyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Department"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Department
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Department");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Department", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.FirstName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string FirstName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FirstName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.JobTitle"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string JobTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "JobTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "JobTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.LastName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string LastName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LastName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.MobileTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string MobileTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MobileTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MobileTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.OfficeLocation"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string OfficeLocation
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OfficeLocation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OfficeLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.PostalCode"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string PostalCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PostalCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.PrimarySmtpAddress"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string PrimarySmtpAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrimarySmtpAddress");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.StateOrProvince"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string StateOrProvince
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StateOrProvince");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StateOrProvince", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.StreetAddress"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string StreetAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StreetAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StreetAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.YomiCompanyName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string YomiCompanyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiCompanyName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiCompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.YomiFirstName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string YomiFirstName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiFirstName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiFirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.YomiLastName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string YomiLastName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiLastName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiLastName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.YomiDisplayName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string YomiDisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiDisplayName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiDisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.YomiDepartment"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string YomiDepartment
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiDepartment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiDepartment", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Delete"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Details"/> </remarks>
		/// <param name="hWnd">optional object hWnd</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Details(object hWnd)
		{
			 Factory.ExecuteMethod(this, "Details", hWnd);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Details"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Details()
		{
			 Factory.ExecuteMethod(this, "Details");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetFreeBusy"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetFreeBusy"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Update"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Update"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.Update"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetContact"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._ContactItem GetContact()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._ContactItem>(this, "GetContact");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetExchangeUser"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ExchangeUser GetExchangeUser()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUser", NetOffice.OutlookApi.ExchangeUser.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetExchangeDistributionList"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ExchangeDistributionList GetExchangeDistributionList()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeDistributionList>(this, "GetExchangeDistributionList", NetOffice.OutlookApi.ExchangeDistributionList.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetDirectReports"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntries GetDirectReports()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetDirectReports", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetMemberOfList"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AddressEntries GetMemberOfList()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetMemberOfList", NetOffice.OutlookApi.AddressEntries.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetExchangeUserManager"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ExchangeUser GetExchangeUserManager()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUserManager", NetOffice.OutlookApi.ExchangeUser.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser.GetPicture"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16), NativeResult]
		public stdole.Picture GetPicture()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPicture", paramsArray);
            return returnItem as stdole.Picture;
        }

		#endregion

		#pragma warning restore
	}
}
