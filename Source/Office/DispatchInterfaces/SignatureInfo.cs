using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SignatureInfo 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865566.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface SignatureInfo : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860243.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool ReadOnly { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865010.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string SignatureProvider { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860281.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string SignatureText { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861498.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        stdole.Picture SignatureImage { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860921.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string SignatureComment { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860572.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.ContentVerificationResults ContentVerificationResults { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864945.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.CertificateVerificationResults CertificateVerificationResults { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862453.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool IsValid { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860786.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool IsCertificateExpired { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865218.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool IsCertificateRevoked { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864566.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool IsCertificateUntrusted { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862539.aspx </remarks>
        /// <param name="sigdet">NetOffice.OfficeApi.Enums.SignatureDetail sigdet</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object GetSignatureDetail(NetOffice.OfficeApi.Enums.SignatureDetail sigdet);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865451.aspx </remarks>
        /// <param name="certdet">NetOffice.OfficeApi.Enums.CertificateDetail certdet</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object GetCertificateDetail(NetOffice.OfficeApi.Enums.CertificateDetail certdet);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863087.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowSignatureCertificate(object parentWindow);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863741.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SelectSignatureCertificate(object parentWindow);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863290.aspx </remarks>
        /// <param name="bstrThumbprint">string bstrThumbprint</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void SelectCertificateDetailByThumbprint(string bstrThumbprint);

        #endregion
    }
}
