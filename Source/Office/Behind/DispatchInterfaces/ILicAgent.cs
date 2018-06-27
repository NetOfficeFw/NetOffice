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
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.ILicAgent);
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
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Initialize", dwBPC, dwMode, bstrLicSource);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetFirstName()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetFirstName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetFirstName(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetFirstName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetLastName()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetLastName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetLastName(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetLastName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetOrgName()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetOrgName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetOrgName(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetOrgName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetEmail()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetEmail");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetEmail(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetEmail", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPhone()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetPhone");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPhone(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPhone", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress1()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress1");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetAddress1(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetAddress1", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCity()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCity");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCity(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCity", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetState()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetState");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetState(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetState", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCountryCode()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCountryCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCountryCode(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCountryCode", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCountryDesc()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCountryDesc");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCountryDesc(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCountryDesc", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetZip()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetZip");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetZip(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetZip", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetIsoLanguage()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetIsoLanguage");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwNewVal">Int32 dwNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetIsoLanguage(Int32 dwNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetIsoLanguage", dwNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetMSUpdate()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetMSUpdate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetMSUpdate(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetMSUpdate", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetMSOffer()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetMSOffer");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetMSOffer(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetMSOffer", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetOtherOffer()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetOtherOffer");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetOtherOffer(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetOtherOffer", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress2()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress2");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetAddress2(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetAddress2", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSystemClock()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSystemClock");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetExistingExpiryDate()
        {
            return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "GetExistingExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetNewExpiryDate()
        {
            return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "GetNewExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingFirstName()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingFirstName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingFirstName(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingFirstName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingLastName()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingLastName");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingLastName(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingLastName", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingPhone()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingPhone");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingPhone(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingPhone", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingAddress1()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingAddress1");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingAddress1(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingAddress1", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingAddress2()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingAddress2");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingAddress2(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingAddress2", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingCity()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingCity");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingCity(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingCity", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingState()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingState");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingState(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingState", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingCountryCode()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingCountryCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingCountryCode(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingCountryCode", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBillingZip()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBillingZip");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetBillingZip(string bstrNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBillingZip", bstrNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bSave">Int32 bSave</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveBillingInfo(Int32 bSave)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveBillingInfo", bSave);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 IsCCRenewalCountry(string bstrCountryCode)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsCCRenewalCountry", bstrCountryCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetVATLabel(string bstrCountryCode)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetVATLabel", bstrCountryCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetCCRenewalExpiryDate()
        {
            return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "GetCCRenewalExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVATNumber">string bstrVATNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetVATNumber(string bstrVATNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetVATNumber", bstrVATNumber);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCCode">string bstrCCCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardType(string bstrCCCode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCreditCardType", bstrCCCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCNumber">string bstrCCNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardNumber(string bstrCCNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCreditCardNumber", bstrCCNumber);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCYear">Int32 dwCCYear</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardExpiryYear(Int32 dwCCYear)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCreditCardExpiryYear", dwCCYear);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCMonth">Int32 dwCCMonth</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCreditCardExpiryMonth(Int32 dwCCMonth)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCreditCardExpiryMonth", dwCCMonth);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardCount()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetCreditCardCount");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardCode(Int32 dwIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCreditCardCode", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardName(Int32 dwIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCreditCardName", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetVATNumber()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetVATNumber");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardType()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCreditCardType");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCreditCardNumber()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCreditCardNumber");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardExpiryYear()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetCreditCardExpiryYear");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCreditCardExpiryMonth()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetCreditCardExpiryMonth");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetDisconnectOption()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetDisconnectOption");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bNewVal">Int32 bNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetDisconnectOption(Int32 bNewVal)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisconnectOption", bNewVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bReviseCustInfo">Int32 bReviseCustInfo</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessHandshakeRequest(Int32 bReviseCustInfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessHandshakeRequest", bReviseCustInfo);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessNewLicenseRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessNewLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessReissueLicenseRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessReissueLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessRetailRenewalLicenseRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessRetailRenewalLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessReviseCustInfoRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessReviseCustInfoRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessCCRenewalPriceRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessCCRenewalPriceRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessCCRenewalLicenseRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessCCRenewalLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetAsyncProcessReturnCode()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetAsyncProcessReturnCode");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 IsUpgradeAvailable()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsUpgradeAvailable");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bWantUpgrade">Int32 bWantUpgrade</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void WantUpgrade(Int32 bWantUpgrade)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "WantUpgrade", bWantUpgrade);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void AsyncProcessDroppedLicenseRequest()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AsyncProcessDroppedLicenseRequest");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GenerateInstallationId()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GenerateInstallationId");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVal">string bstrVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DepositConfirmationId(string bstrVal)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DepositConfirmationId", bstrVal);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCIDIID">string bstrCIDIID</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 VerifyCheckDigits(string bstrCIDIID)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "VerifyCheckDigits", bstrCIDIID);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual DateTime GetCurrentExpiryDate()
        {
            return InvokerService.InvokeInternal.ExecuteDateTimeMethodGet(this, "GetCurrentExpiryDate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bIsLicenseRequest">Int32 bIsLicenseRequest</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void CancelAsyncProcessRequest(Int32 bIsLicenseRequest)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CancelAsyncProcessRequest", bIsLicenseRequest);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetCurrencyDescription(Int32 dwCurrencyIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCurrencyDescription", dwCurrencyIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetPriceItemCount()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetPriceItemCount");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPriceItemLabel(Int32 dwIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetPriceItemLabel", dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetPriceItemValue(Int32 dwCurrencyIndex, Int32 dwIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetPriceItemValue", dwCurrencyIndex, dwIndex);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetInvoiceText()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetInvoiceText");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetBackendErrorMsg()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetBackendErrorMsg");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetCurrencyOption()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetCurrencyOption");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyOption">Int32 dwCurrencyOption</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetCurrencyOption(Int32 dwCurrencyOption)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCurrencyOption", dwCurrencyOption);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string GetEndOfLifeHtmlText()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetEndOfLifeHtmlText");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DisplaySSLCert()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DisplaySSLCert");
        }

        #endregion

        #pragma warning restore
    }
}
