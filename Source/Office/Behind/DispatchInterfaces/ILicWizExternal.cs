using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ILicWizExternal 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ILicWizExternal : COMObject, NetOffice.OfficeApi.ILicWizExternal
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
                    _contractType = typeof(NetOffice.OfficeApi.ILicWizExternal);
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
                    _type = typeof(ILicWizExternal);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ILicWizExternal() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Context
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Context");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Validator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Validator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object LicAgent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "LicAgent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string CountryInfo
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CountryInfo");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 WizardVisible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WizardVisible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WizardVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string WizardTitle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WizardTitle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WizardTitle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 AnimationEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AnimationEnabled");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CurrentHelpId
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentHelpId");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentHelpId", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string OfficeOnTheWebUrl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OfficeOnTheWebUrl");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="punkHtmlDoc">object punkHtmlDoc</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintHtmlDocument(object punkHtmlDoc)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintHtmlDocument", punkHtmlDoc);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void InvokeDateTimeApplet()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InvokeDateTimeApplet");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="date">DateTime date</param>
        /// <param name="pFormat">optional string pFormat = </param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string FormatDate(DateTime date, object pFormat)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatDate", date, pFormat);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="date">DateTime date</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string FormatDate(DateTime date)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "FormatDate", date);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pvarId">optional object pvarId</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ShowHelp(object pvarId)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp", pvarId);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ShowHelp()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Terminate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Terminate");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bPC">Int32 bPC</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void DisableVORWReminder(Int32 bPC)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DisableVORWReminder", bPC);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrReceipt">string bstrReceipt</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string SaveReceipt(string bstrReceipt)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "SaveReceipt", bstrReceipt);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrUrl">string bstrUrl</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void OpenInDefaultBrowser(string bstrUrl)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OpenInDefaultBrowser", bstrUrl);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrButtons">string bstrButtons</param>
        /// <param name="bstrIcon">string bstrIcon</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MsoAlert(string bstrText, string bstrButtons, string bstrIcon)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MsoAlert", bstrText, bstrButtons, bstrIcon);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrKey">string bstrKey</param>
        /// <param name="fMORW">Int32 fMORW</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DepositPidKey(string bstrKey, Int32 fMORW)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DepositPidKey", bstrKey, fMORW);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrMessage">string bstrMessage</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void WriteLog(string bstrMessage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "WriteLog", bstrMessage);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrProductCode">string bstrProductCode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ResignDpc(string bstrProductCode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResignDpc", bstrProductCode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ResetPID()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetPID");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dx">Int32 dx</param>
        /// <param name="dy">Int32 dy</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetDialogSize(Int32 dx, Int32 dy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDialogSize", dx, dy);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lMode">Int32 lMode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 VerifyClock(Int32 lMode)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "VerifyClock", lMode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pdispSelect">object pdispSelect</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SortSelectOptions(object pdispSelect)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortSelectOptions", pdispSelect);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void InternetDisconnect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InternetDisconnect");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GetConnectedState()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetConnectedState");
        }

        #endregion

        #pragma warning restore
    }
}
