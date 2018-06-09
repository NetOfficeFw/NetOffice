using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface ILicWizExternal 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("4CAC6328-B9B0-11D3-8D59-0050048384E3")]
    public interface ILicWizExternal : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Context { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Validator { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object LicAgent { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string CountryInfo { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 WizardVisible { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string WizardTitle { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 AnimationEnabled { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 CurrentHelpId { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string OfficeOnTheWebUrl { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="punkHtmlDoc">object punkHtmlDoc</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void PrintHtmlDocument(object punkHtmlDoc);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void InvokeDateTimeApplet();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="date">DateTime date</param>
        /// <param name="pFormat">optional string pFormat = </param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string FormatDate(DateTime date, object pFormat);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="date">DateTime date</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string FormatDate(DateTime date);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pvarId">optional object pvarId</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void ShowHelp(object pvarId);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void ShowHelp();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Terminate();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bPC">Int32 bPC</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void DisableVORWReminder(Int32 bPC);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrReceipt">string bstrReceipt</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string SaveReceipt(string bstrReceipt);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrUrl">string bstrUrl</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void OpenInDefaultBrowser(string bstrUrl);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrButtons">string bstrButtons</param>
        /// <param name="bstrIcon">string bstrIcon</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 MsoAlert(string bstrText, string bstrButtons, string bstrIcon);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrKey">string bstrKey</param>
        /// <param name="fMORW">Int32 fMORW</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 DepositPidKey(string bstrKey, Int32 fMORW);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrMessage">string bstrMessage</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void WriteLog(string bstrMessage);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrProductCode">string bstrProductCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void ResignDpc(string bstrProductCode);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void ResetPID();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dx">Int32 dx</param>
        /// <param name="dy">Int32 dy</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetDialogSize(Int32 dx, Int32 dy);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lMode">Int32 lMode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 VerifyClock(Int32 lMode);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pdispSelect">object pdispSelect</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SortSelectOptions(object pdispSelect);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void InternetDisconnect();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 GetConnectedState();

        #endregion
    }
}
