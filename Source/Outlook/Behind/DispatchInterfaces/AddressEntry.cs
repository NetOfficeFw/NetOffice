using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface AddressEntry 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869306.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AddressEntry : COMObject, NetOffice.OutlookApi.AddressEntry
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
                    _contractType = typeof(NetOffice.OutlookApi.AddressEntry);
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
                    _type = typeof(AddressEntry);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public AddressEntry() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866732.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865346.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869587.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861815.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863620.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869330.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlDisplayType DisplayType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDisplayType>(this, "DisplayType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860641.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.AddressEntry Manager
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntry>(this, "Manager", typeof(NetOffice.OutlookApi.AddressEntry));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.AddressEntries Members
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AddressEntries>(this, "Members", typeof(NetOffice.OutlookApi.AddressEntries));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863077.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862424.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860674.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866395.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", typeof(NetOffice.OutlookApi.PropertyAccessor));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865383.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867301.aspx </remarks>
		/// <param name="hWnd">optional object hWnd</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Details(object hWnd)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Details", hWnd);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867301.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Details()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Details");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867624.aspx </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		/// <param name="completeFormat">optional object completeFormat</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string GetFreeBusy(DateTime start, Int32 minPerChar, object completeFormat)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar, completeFormat);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867624.aspx </remarks>
		/// <param name="start">DateTime start</param>
		/// <param name="minPerChar">Int32 minPerChar</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string GetFreeBusy(DateTime start, Int32 minPerChar)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetFreeBusy", start, minPerChar);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860722.aspx </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		/// <param name="refresh">optional object refresh</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Update(object makePermanent, object refresh)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", makePermanent, refresh);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860722.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Update()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860722.aspx </remarks>
		/// <param name="makePermanent">optional object makePermanent</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Update(object makePermanent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", makePermanent);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void UpdateFreeBusy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateFreeBusy");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862409.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._ContactItem GetContact()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._ContactItem>(this, "GetContact");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869721.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.ExchangeUser GetExchangeUser()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeUser>(this, "GetExchangeUser", typeof(NetOffice.OutlookApi.ExchangeUser));
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860629.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual NetOffice.OutlookApi.ExchangeDistributionList GetExchangeDistributionList()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.ExchangeDistributionList>(this, "GetExchangeDistributionList", typeof(NetOffice.OutlookApi.ExchangeDistributionList));
		}

		#endregion

		#pragma warning restore
	}
}


