using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _ExchangeUser 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _ExchangeUser : COMObject, NetOffice.OutlookApi._ExchangeUser
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
                    _contractType = typeof(NetOffice.OutlookApi._ExchangeUser);
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
                    _type = typeof(_ExchangeUser);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _ExchangeUser() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861610.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869804.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866970.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861801.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868642.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Address
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863069.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlDisplayType DisplayType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDisplayType>(this, "DisplayType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868616.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OutlookApi.AddressEntry Manager
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntry>(this, "Manager", typeof(NetOffice.OutlookApi.AddressEntry));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object MAPIOBJECT
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "MAPIOBJECT", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OutlookApi.AddressEntries Members
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntries>(this, "Members", typeof(NetOffice.OutlookApi.AddressEntries));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867518.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869493.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870115.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlAddressEntryUserType AddressEntryUserType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlAddressEntryUserType>(this, "AddressEntryUserType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869240.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869719.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Alias
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Alias");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869132.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string AssistantName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AssistantName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AssistantName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868874.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string BusinessTelephoneNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BusinessTelephoneNumber");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BusinessTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870145.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string City
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "City");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "City", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868681.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Comments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Comments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Comments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869367.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string CompanyName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CompanyName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863675.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string Department
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Department");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Department", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866187.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string FirstName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FirstName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862975.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string JobTitle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "JobTitle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "JobTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862155.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string LastName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LastName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LastName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868169.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string MobileTelephoneNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MobileTelephoneNumber");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MobileTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868640.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string OfficeLocation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OfficeLocation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OfficeLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868582.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string PostalCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PostalCode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862991.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string PrimarySmtpAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PrimarySmtpAddress");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868486.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string StateOrProvince
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StateOrProvince");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StateOrProvince", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861563.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string StreetAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StreetAddress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StreetAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864425.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string YomiCompanyName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "YomiCompanyName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YomiCompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868654.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string YomiFirstName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "YomiFirstName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YomiFirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860660.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string YomiLastName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "YomiLastName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YomiLastName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866433.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string YomiDisplayName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "YomiDisplayName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YomiDisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866216.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string YomiDepartment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "YomiDepartment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YomiDepartment", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869237.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866234.aspx </remarks>
		/// <param name="hWnd">optional object hWnd</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Details(object hWnd)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Details", hWnd);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866234.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Details()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Details");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860983.aspx </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		/// <param name="completeFormat">optional object completeFormat</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string GetFreeBusy(DateTime start, Int32 minPerChar, object completeFormat)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar, completeFormat);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860983.aspx </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual string GetFreeBusy(DateTime start, Int32 minPerChar)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868285.aspx </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		/// <param name="refresh">optional object refresh</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Update(object makePermanent, object refresh)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", makePermanent, refresh);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868285.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Update()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868285.aspx </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Update(object makePermanent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", makePermanent);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void UpdateFreeBusy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateFreeBusy");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864255.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._ContactItem GetContact()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._ContactItem>(this, "GetContact");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870184.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.ExchangeUser GetExchangeUser()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUser", typeof(NetOffice.OutlookApi.ExchangeUser));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864766.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.ExchangeDistributionList GetExchangeDistributionList()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeDistributionList>(this, "GetExchangeDistributionList", typeof(NetOffice.OutlookApi.ExchangeDistributionList));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866704.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.AddressEntries GetDirectReports()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetDirectReports", typeof(NetOffice.OutlookApi.AddressEntries));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862143.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.AddressEntries GetMemberOfList()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntries>(this, "GetMemberOfList", typeof(NetOffice.OutlookApi.AddressEntries));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869724.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.ExchangeUser GetExchangeUserManager()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUserManager", typeof(NetOffice.OutlookApi.ExchangeUser));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864210.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), NativeResult]
		public virtual stdole.Picture GetPicture()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPicture", paramsArray);
            return returnItem as stdole.Picture;
        }

		#endregion

		#pragma warning restore
	}
}


