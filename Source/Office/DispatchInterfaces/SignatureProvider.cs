using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SignatureProvider 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861225.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000CD6A3-0000-0000-C000-000000000046")]
    public interface SignatureProvider : ICOMObject
    {
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
        stdole.Picture GenerateSignatureLineImage(NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861424.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowSignatureSetup(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864670.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowSigningCeremony(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864683.aspx </remarks>
        /// <param name="queryContinue">object queryContinue</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        /// <param name="xmlDsigStream">object xmlDsigStream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SignXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860266.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
        /// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void NotifySignatureAdded(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo);

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
        void VerifyXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres);

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
        void ShowSignatureDetails(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863646.aspx </remarks>
        /// <param name="sigprovdet">NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object GetProviderDetail(NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862104.aspx </remarks>
        /// <param name="queryContinue">object queryContinue</param>
        /// <param name="stream">object stream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        byte[] HashStream(object queryContinue, object stream);

        #endregion
    }
}
