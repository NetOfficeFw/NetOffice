using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface ILicAgent 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ILicAgent : COMObject
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
                    _type = typeof(ILicAgent);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ILicAgent(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ILicAgent(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwBPC">Int32 dwBPC</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="bstrLicSource">string bstrLicSource</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 Initialize(Int32 dwBPC, Int32 dwMode, string bstrLicSource)
		{
			return Factory.ExecuteInt32MethodGet(this, "Initialize", dwBPC, dwMode, bstrLicSource);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetFirstName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetFirstName");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetFirstName(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetFirstName", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetLastName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetLastName");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetLastName(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetLastName", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetOrgName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetOrgName");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetOrgName(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetOrgName", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetEmail()
		{
			return Factory.ExecuteStringMethodGet(this, "GetEmail");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetEmail(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetEmail", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetPhone()
		{
			return Factory.ExecuteStringMethodGet(this, "GetPhone");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetPhone(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetPhone", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetAddress1()
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress1");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetAddress1(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetAddress1", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCity()
		{
			return Factory.ExecuteStringMethodGet(this, "GetCity");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCity(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetCity", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetState()
		{
			return Factory.ExecuteStringMethodGet(this, "GetState");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetState(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetState", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCountryCode()
		{
			return Factory.ExecuteStringMethodGet(this, "GetCountryCode");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCountryCode(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetCountryCode", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCountryDesc()
		{
			return Factory.ExecuteStringMethodGet(this, "GetCountryDesc");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCountryDesc(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetCountryDesc", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetZip()
		{
			return Factory.ExecuteStringMethodGet(this, "GetZip");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetZip(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetZip", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetIsoLanguage()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetIsoLanguage");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwNewVal">Int32 dwNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetIsoLanguage(Int32 dwNewVal)
		{
			 Factory.ExecuteMethod(this, "SetIsoLanguage", dwNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetMSUpdate()
		{
			return Factory.ExecuteStringMethodGet(this, "GetMSUpdate");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetMSUpdate(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetMSUpdate", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetMSOffer()
		{
			return Factory.ExecuteStringMethodGet(this, "GetMSOffer");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetMSOffer(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetMSOffer", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetOtherOffer()
		{
			return Factory.ExecuteStringMethodGet(this, "GetOtherOffer");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetOtherOffer(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetOtherOffer", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetAddress2()
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress2");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetAddress2(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetAddress2", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 CheckSystemClock()
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSystemClock");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public DateTime GetExistingExpiryDate()
		{
			return Factory.ExecuteDateTimeMethodGet(this, "GetExistingExpiryDate");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public DateTime GetNewExpiryDate()
		{
			return Factory.ExecuteDateTimeMethodGet(this, "GetNewExpiryDate");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingFirstName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingFirstName");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingFirstName(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingFirstName", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingLastName()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingLastName");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingLastName(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingLastName", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingPhone()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingPhone");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingPhone(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingPhone", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingAddress1()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingAddress1");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingAddress1(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingAddress1", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingAddress2()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingAddress2");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingAddress2(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingAddress2", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingCity()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingCity");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingCity(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingCity", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingState()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingState");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingState(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingState", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingCountryCode()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingCountryCode");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingCountryCode(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingCountryCode", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBillingZip()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBillingZip");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetBillingZip(string bstrNewVal)
		{
			 Factory.ExecuteMethod(this, "SetBillingZip", bstrNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bSave">Int32 bSave</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 SaveBillingInfo(Int32 bSave)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveBillingInfo", bSave);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCountryCode">string bstrCountryCode</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 IsCCRenewalCountry(string bstrCountryCode)
		{
			return Factory.ExecuteInt32MethodGet(this, "IsCCRenewalCountry", bstrCountryCode);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCountryCode">string bstrCountryCode</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetVATLabel(string bstrCountryCode)
		{
			return Factory.ExecuteStringMethodGet(this, "GetVATLabel", bstrCountryCode);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public DateTime GetCCRenewalExpiryDate()
		{
			return Factory.ExecuteDateTimeMethodGet(this, "GetCCRenewalExpiryDate");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrVATNumber">string bstrVATNumber</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetVATNumber(string bstrVATNumber)
		{
			 Factory.ExecuteMethod(this, "SetVATNumber", bstrVATNumber);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCCCode">string bstrCCCode</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCreditCardType(string bstrCCCode)
		{
			 Factory.ExecuteMethod(this, "SetCreditCardType", bstrCCCode);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCCNumber">string bstrCCNumber</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCreditCardNumber(string bstrCCNumber)
		{
			 Factory.ExecuteMethod(this, "SetCreditCardNumber", bstrCCNumber);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwCCYear">Int32 dwCCYear</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCreditCardExpiryYear(Int32 dwCCYear)
		{
			 Factory.ExecuteMethod(this, "SetCreditCardExpiryYear", dwCCYear);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwCCMonth">Int32 dwCCMonth</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCreditCardExpiryMonth(Int32 dwCCMonth)
		{
			 Factory.ExecuteMethod(this, "SetCreditCardExpiryMonth", dwCCMonth);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetCreditCardCount");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCreditCardCode(Int32 dwIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCreditCardCode", dwIndex);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCreditCardName(Int32 dwIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCreditCardName", dwIndex);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetVATNumber()
		{
			return Factory.ExecuteStringMethodGet(this, "GetVATNumber");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCreditCardType()
		{
			return Factory.ExecuteStringMethodGet(this, "GetCreditCardType");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCreditCardNumber()
		{
			return Factory.ExecuteStringMethodGet(this, "GetCreditCardNumber");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardExpiryYear()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetCreditCardExpiryYear");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardExpiryMonth()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetCreditCardExpiryMonth");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetDisconnectOption()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetDisconnectOption");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bNewVal">Int32 bNewVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetDisconnectOption(Int32 bNewVal)
		{
			 Factory.ExecuteMethod(this, "SetDisconnectOption", bNewVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bReviseCustInfo">Int32 bReviseCustInfo</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessHandshakeRequest(Int32 bReviseCustInfo)
		{
			 Factory.ExecuteMethod(this, "AsyncProcessHandshakeRequest", bReviseCustInfo);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessNewLicenseRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessNewLicenseRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessReissueLicenseRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessReissueLicenseRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessRetailRenewalLicenseRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessRetailRenewalLicenseRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessReviseCustInfoRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessReviseCustInfoRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessCCRenewalPriceRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessCCRenewalPriceRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessCCRenewalLicenseRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessCCRenewalLicenseRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetAsyncProcessReturnCode()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetAsyncProcessReturnCode");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 IsUpgradeAvailable()
		{
			return Factory.ExecuteInt32MethodGet(this, "IsUpgradeAvailable");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bWantUpgrade">Int32 bWantUpgrade</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void WantUpgrade(Int32 bWantUpgrade)
		{
			 Factory.ExecuteMethod(this, "WantUpgrade", bWantUpgrade);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void AsyncProcessDroppedLicenseRequest()
		{
			 Factory.ExecuteMethod(this, "AsyncProcessDroppedLicenseRequest");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GenerateInstallationId()
		{
			return Factory.ExecuteStringMethodGet(this, "GenerateInstallationId");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrVal">string bstrVal</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 DepositConfirmationId(string bstrVal)
		{
			return Factory.ExecuteInt32MethodGet(this, "DepositConfirmationId", bstrVal);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCIDIID">string bstrCIDIID</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 VerifyCheckDigits(string bstrCIDIID)
		{
			return Factory.ExecuteInt32MethodGet(this, "VerifyCheckDigits", bstrCIDIID);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public DateTime GetCurrentExpiryDate()
		{
			return Factory.ExecuteDateTimeMethodGet(this, "GetCurrentExpiryDate");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bIsLicenseRequest">Int32 bIsLicenseRequest</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void CancelAsyncProcessRequest(Int32 bIsLicenseRequest)
		{
			 Factory.ExecuteMethod(this, "CancelAsyncProcessRequest", bIsLicenseRequest);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetCurrencyDescription(Int32 dwCurrencyIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCurrencyDescription", dwCurrencyIndex);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetPriceItemCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetPriceItemCount");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetPriceItemLabel(Int32 dwIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetPriceItemLabel", dwIndex);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetPriceItemValue(Int32 dwCurrencyIndex, Int32 dwIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetPriceItemValue", dwCurrencyIndex, dwIndex);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetInvoiceText()
		{
			return Factory.ExecuteStringMethodGet(this, "GetInvoiceText");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetBackendErrorMsg()
		{
			return Factory.ExecuteStringMethodGet(this, "GetBackendErrorMsg");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 GetCurrencyOption()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetCurrencyOption");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dwCurrencyOption">Int32 dwCurrencyOption</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetCurrencyOption(Int32 dwCurrencyOption)
		{
			 Factory.ExecuteMethod(this, "SetCurrencyOption", dwCurrencyOption);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string GetEndOfLifeHtmlText()
		{
			return Factory.ExecuteStringMethodGet(this, "GetEndOfLifeHtmlText");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 DisplaySSLCert()
		{
			return Factory.ExecuteInt32MethodGet(this, "DisplaySSLCert");
		}

		#endregion

		#pragma warning restore
	}
}
