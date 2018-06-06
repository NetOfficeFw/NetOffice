using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ILicAgent 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ILicAgent : COMObject, NetOffice.OfficeApi.ILicAgent
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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ILicAgent() : base()
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Initialize(Int32 dwBPC, Int32 dwMode, string bstrLicSource)
        {
            return Factory.ExecuteInt32MethodGet(this, "Initialize", dwBPC, dwMode, bstrLicSource);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetFirstName()
        {
            return Factory.ExecuteStringMethodGet(this, "GetFirstName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetFirstName(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetFirstName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetLastName()
        {
            return Factory.ExecuteStringMethodGet(this, "GetLastName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetLastName(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetLastName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetOrgName()
        {
            return Factory.ExecuteStringMethodGet(this, "GetOrgName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetOrgName(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetOrgName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetEmail()
        {
            return Factory.ExecuteStringMethodGet(this, "GetEmail");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetEmail(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetEmail", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPhone()
        {
            return Factory.ExecuteStringMethodGet(this, "GetPhone");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPhone(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetPhone", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress1()
        {
            return Factory.ExecuteStringMethodGet(this, "GetAddress1");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetAddress1(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetAddress1", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCity()
        {
            return Factory.ExecuteStringMethodGet(this, "GetCity");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCity(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetCity", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetState()
        {
            return Factory.ExecuteStringMethodGet(this, "GetState");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetState(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetState", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCountryCode()
        {
            return Factory.ExecuteStringMethodGet(this, "GetCountryCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCountryCode(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetCountryCode", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCountryDesc()
        {
            return Factory.ExecuteStringMethodGet(this, "GetCountryDesc");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCountryDesc(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetCountryDesc", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetZip()
        {
            return Factory.ExecuteStringMethodGet(this, "GetZip");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetZip(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetZip", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetIsoLanguage()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetIsoLanguage");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwNewVal">Int32 dwNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetIsoLanguage(Int32 dwNewVal)
        {
            Factory.ExecuteMethod(this, "SetIsoLanguage", dwNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetMSUpdate()
        {
            return Factory.ExecuteStringMethodGet(this, "GetMSUpdate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetMSUpdate(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetMSUpdate", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetMSOffer()
        {
            return Factory.ExecuteStringMethodGet(this, "GetMSOffer");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetMSOffer(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetMSOffer", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetOtherOffer()
        {
            return Factory.ExecuteStringMethodGet(this, "GetOtherOffer");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetOtherOffer(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetOtherOffer", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress2()
        {
            return Factory.ExecuteStringMethodGet(this, "GetAddress2");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetAddress2(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetAddress2", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSystemClock()
        {
            return Factory.ExecuteInt32MethodGet(this, "CheckSystemClock");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetExistingExpiryDate()
        {
            return Factory.ExecuteDateTimeMethodGet(this, "GetExistingExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetNewExpiryDate()
        {
            return Factory.ExecuteDateTimeMethodGet(this, "GetNewExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingFirstName()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingFirstName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingFirstName(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingFirstName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingLastName()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingLastName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingLastName(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingLastName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingPhone()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingPhone");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingPhone(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingPhone", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingAddress1()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingAddress1");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingAddress1(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingAddress1", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingAddress2()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingAddress2");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingAddress2(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingAddress2", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingCity()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingCity");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingCity(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingCity", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingState()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingState");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingState(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingState", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingCountryCode()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingCountryCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingCountryCode(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingCountryCode", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingZip()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBillingZip");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingZip(string bstrNewVal)
        {
            Factory.ExecuteMethod(this, "SetBillingZip", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bSave">Int32 bSave</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveBillingInfo(Int32 bSave)
        {
            return Factory.ExecuteInt32MethodGet(this, "SaveBillingInfo", bSave);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 IsCCRenewalCountry(string bstrCountryCode)
        {
            return Factory.ExecuteInt32MethodGet(this, "IsCCRenewalCountry", bstrCountryCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetVATLabel(string bstrCountryCode)
        {
            return Factory.ExecuteStringMethodGet(this, "GetVATLabel", bstrCountryCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetCCRenewalExpiryDate()
        {
            return Factory.ExecuteDateTimeMethodGet(this, "GetCCRenewalExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVATNumber">string bstrVATNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetVATNumber(string bstrVATNumber)
        {
            Factory.ExecuteMethod(this, "SetVATNumber", bstrVATNumber);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCCode">string bstrCCCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardType(string bstrCCCode)
        {
            Factory.ExecuteMethod(this, "SetCreditCardType", bstrCCCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCNumber">string bstrCCNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardNumber(string bstrCCNumber)
        {
            Factory.ExecuteMethod(this, "SetCreditCardNumber", bstrCCNumber);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCYear">Int32 dwCCYear</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardExpiryYear(Int32 dwCCYear)
        {
            Factory.ExecuteMethod(this, "SetCreditCardExpiryYear", dwCCYear);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCMonth">Int32 dwCCMonth</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardExpiryMonth(Int32 dwCCMonth)
        {
            Factory.ExecuteMethod(this, "SetCreditCardExpiryMonth", dwCCMonth);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardCount()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetCreditCardCount");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardCode(Int32 dwIndex)
        {
            return Factory.ExecuteStringMethodGet(this, "GetCreditCardCode", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardName(Int32 dwIndex)
        {
            return Factory.ExecuteStringMethodGet(this, "GetCreditCardName", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetVATNumber()
        {
            return Factory.ExecuteStringMethodGet(this, "GetVATNumber");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardType()
        {
            return Factory.ExecuteStringMethodGet(this, "GetCreditCardType");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardNumber()
        {
            return Factory.ExecuteStringMethodGet(this, "GetCreditCardNumber");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardExpiryYear()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetCreditCardExpiryYear");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardExpiryMonth()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetCreditCardExpiryMonth");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetDisconnectOption()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetDisconnectOption");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bNewVal">Int32 bNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetDisconnectOption(Int32 bNewVal)
        {
            Factory.ExecuteMethod(this, "SetDisconnectOption", bNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bReviseCustInfo">Int32 bReviseCustInfo</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessHandshakeRequest(Int32 bReviseCustInfo)
        {
            Factory.ExecuteMethod(this, "AsyncProcessHandshakeRequest", bReviseCustInfo);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessNewLicenseRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessNewLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessReissueLicenseRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessReissueLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessRetailRenewalLicenseRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessRetailRenewalLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessReviseCustInfoRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessReviseCustInfoRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessCCRenewalPriceRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessCCRenewalPriceRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessCCRenewalLicenseRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessCCRenewalLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetAsyncProcessReturnCode()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetAsyncProcessReturnCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 IsUpgradeAvailable()
        {
            return Factory.ExecuteInt32MethodGet(this, "IsUpgradeAvailable");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bWantUpgrade">Int32 bWantUpgrade</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void WantUpgrade(Int32 bWantUpgrade)
        {
            Factory.ExecuteMethod(this, "WantUpgrade", bWantUpgrade);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessDroppedLicenseRequest()
        {
            Factory.ExecuteMethod(this, "AsyncProcessDroppedLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GenerateInstallationId()
        {
            return Factory.ExecuteStringMethodGet(this, "GenerateInstallationId");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVal">string bstrVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DepositConfirmationId(string bstrVal)
        {
            return Factory.ExecuteInt32MethodGet(this, "DepositConfirmationId", bstrVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCIDIID">string bstrCIDIID</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 VerifyCheckDigits(string bstrCIDIID)
        {
            return Factory.ExecuteInt32MethodGet(this, "VerifyCheckDigits", bstrCIDIID);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetCurrentExpiryDate()
        {
            return Factory.ExecuteDateTimeMethodGet(this, "GetCurrentExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bIsLicenseRequest">Int32 bIsLicenseRequest</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void CancelAsyncProcessRequest(Int32 bIsLicenseRequest)
        {
            Factory.ExecuteMethod(this, "CancelAsyncProcessRequest", bIsLicenseRequest);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCurrencyDescription(Int32 dwCurrencyIndex)
        {
            return Factory.ExecuteStringMethodGet(this, "GetCurrencyDescription", dwCurrencyIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetPriceItemCount()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetPriceItemCount");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPriceItemLabel(Int32 dwIndex)
        {
            return Factory.ExecuteStringMethodGet(this, "GetPriceItemLabel", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPriceItemValue(Int32 dwCurrencyIndex, Int32 dwIndex)
        {
            return Factory.ExecuteStringMethodGet(this, "GetPriceItemValue", dwCurrencyIndex, dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetInvoiceText()
        {
            return Factory.ExecuteStringMethodGet(this, "GetInvoiceText");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBackendErrorMsg()
        {
            return Factory.ExecuteStringMethodGet(this, "GetBackendErrorMsg");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCurrencyOption()
        {
            return Factory.ExecuteInt32MethodGet(this, "GetCurrencyOption");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyOption">Int32 dwCurrencyOption</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCurrencyOption(Int32 dwCurrencyOption)
        {
            Factory.ExecuteMethod(this, "SetCurrencyOption", dwCurrencyOption);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetEndOfLifeHtmlText()
        {
            return Factory.ExecuteStringMethodGet(this, "GetEndOfLifeHtmlText");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DisplaySSLCert()
        {
            return Factory.ExecuteInt32MethodGet(this, "DisplaySSLCert");
        }

        #endregion

        #pragma warning restore
    }
}
