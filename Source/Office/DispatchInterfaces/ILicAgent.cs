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
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00194002-D9C3-11D3-8D59-0050048384E3")]
    public interface ILicAgent : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwBPC">Int32 dwBPC</param>
        /// <param name="dwMode">Int32 dwMode</param>
        /// <param name="bstrLicSource">string bstrLicSource</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Initialize(Int32 dwBPC, Int32 dwMode, string bstrLicSource);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetFirstName();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetFirstName(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetLastName();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetLastName(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetOrgName();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetOrgName(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetEmail();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetEmail(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetPhone();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetPhone(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetAddress1();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetAddress1(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCity();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCity(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetState();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetState(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCountryCode();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCountryCode(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCountryDesc();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCountryDesc(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetZip();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetZip(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetIsoLanguage();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwNewVal">Int32 dwNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetIsoLanguage(Int32 dwNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetMSUpdate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetMSUpdate(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetMSOffer();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetMSOffer(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetOtherOffer();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetOtherOffer(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetAddress2();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetAddress2(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 CheckSystemClock();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        DateTime GetExistingExpiryDate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        DateTime GetNewExpiryDate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingFirstName();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingFirstName(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingLastName();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingLastName(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingPhone();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingPhone(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingAddress1();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingAddress1(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingAddress2();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingAddress2(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingCity();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingCity(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingState();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingState(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingCountryCode();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingCountryCode(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBillingZip();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrNewVal">string bstrNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetBillingZip(string bstrNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bSave">Int32 bSave</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 SaveBillingInfo(Int32 bSave);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 IsCCRenewalCountry(string bstrCountryCode);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCountryCode">string bstrCountryCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetVATLabel(string bstrCountryCode);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        DateTime GetCCRenewalExpiryDate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVATNumber">string bstrVATNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetVATNumber(string bstrVATNumber);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCCode">string bstrCCCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCreditCardType(string bstrCCCode);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCCNumber">string bstrCCNumber</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCreditCardNumber(string bstrCCNumber);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCYear">Int32 dwCCYear</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCreditCardExpiryYear(Int32 dwCCYear);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCCMonth">Int32 dwCCMonth</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCreditCardExpiryMonth(Int32 dwCCMonth);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetCreditCardCount();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCreditCardCode(Int32 dwIndex);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCreditCardName(Int32 dwIndex);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetVATNumber();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCreditCardType();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCreditCardNumber();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetCreditCardExpiryYear();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetCreditCardExpiryMonth();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetDisconnectOption();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bNewVal">Int32 bNewVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetDisconnectOption(Int32 bNewVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bReviseCustInfo">Int32 bReviseCustInfo</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessHandshakeRequest(Int32 bReviseCustInfo);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessNewLicenseRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessReissueLicenseRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessRetailRenewalLicenseRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessReviseCustInfoRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessCCRenewalPriceRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessCCRenewalLicenseRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetAsyncProcessReturnCode();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 IsUpgradeAvailable();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bWantUpgrade">Int32 bWantUpgrade</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void WantUpgrade(Int32 bWantUpgrade);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void AsyncProcessDroppedLicenseRequest();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GenerateInstallationId();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrVal">string bstrVal</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 DepositConfirmationId(string bstrVal);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrCIDIID">string bstrCIDIID</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 VerifyCheckDigits(string bstrCIDIID);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        DateTime GetCurrentExpiryDate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bIsLicenseRequest">Int32 bIsLicenseRequest</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void CancelAsyncProcessRequest(Int32 bIsLicenseRequest);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetCurrencyDescription(Int32 dwCurrencyIndex);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetPriceItemCount();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetPriceItemLabel(Int32 dwIndex);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
        /// <param name="dwIndex">Int32 dwIndex</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetPriceItemValue(Int32 dwCurrencyIndex, Int32 dwIndex);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetInvoiceText();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetBackendErrorMsg();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetCurrencyOption();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dwCurrencyOption">Int32 dwCurrencyOption</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetCurrencyOption(Int32 dwCurrencyOption);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string GetEndOfLifeHtmlText();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 DisplaySSLCert();

        #endregion
    }
}
