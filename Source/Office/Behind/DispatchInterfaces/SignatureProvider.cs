using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SignatureProvider 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861225.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class SignatureProvider : COMObject, NetOffice.OfficeApi.SignatureProvider
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
                    _contractType = typeof(NetOffice.OfficeApi.SignatureProvider);
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
                    _type = typeof(SignatureProvider);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SignatureProvider() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861157.aspx </remarks>
        /// <param name="siglnimg">NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        /// <param name="xmlDsigStream">object xmlDsigStream</param>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        public virtual stdole.Picture GenerateSignatureLineImage(NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
        {
            object[] paramsArray = new object[] { siglnimg, psigsetup, psiginfo, xmlDsigStream };
            object returnItem = InvokerService.InvokeInternal.ExecuteObjectMethodGet(this, "GenerateSignatureLineImage", paramsArray); ;            
            return returnItem as stdole.Picture;
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861424.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowSignatureSetup(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowSignatureSetup", parentWindow, psigsetup);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864670.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowSigningCeremony(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowSigningCeremony", parentWindow, psigsetup, psiginfo);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864683.aspx </remarks>
        /// <param name="queryContinue">object queryContinue</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        /// <param name="xmlDsigStream">object xmlDsigStream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SignXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SignXmlDsig", queryContinue, psigsetup, psiginfo, xmlDsigStream);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860266.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void NotifySignatureAdded(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NotifySignatureAdded", parentWindow, psigsetup, psiginfo);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863028.aspx </remarks>
        /// <param name="queryContinue">object queryContinue</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        /// <param name="xmlDsigStream">object xmlDsigStream</param>
        /// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
        /// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void VerifyXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "VerifyXmlDsig", new object[] { queryContinue, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865248.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        /// <param name="xmlDsigStream">object xmlDsigStream</param>
        /// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
        /// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowSignatureDetails(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowSignatureDetails", new object[] { parentWindow, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863646.aspx </remarks>
        /// <param name="sigprovdet">NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object GetProviderDetail(NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetProviderDetail", sigprovdet);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862104.aspx </remarks>
        /// <param name="queryContinue">object queryContinue</param>
        /// <param name="stream">object stream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual byte[] HashStream(object queryContinue, object stream)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(queryContinue, stream);
            object returnItem = (object)Invoker.MethodReturn(this, "HashStream", paramsArray);
            return (byte[])returnItem;
        }

        #endregion

        #pragma warning restore
    }
}
