using stdole;
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// Represents a signature provider add-in.
    /// </summary>
    /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.aspx </remarks>
    [ComImport, Guid("000CD6A3-0000-0000-C000-000000000046"), TypeLibType(4160)]
    public interface SignatureProvider
    {
        /// <summary>
        /// Gets a signature line image.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.generatesignaturelineimage.aspx </remarks>
        /// <param name="siglnimg">Contains the name if the signature line graphic.</param>
        /// <param name="psigsetup">Specifies initial settings of the signature provider add-in.</param>
        /// <param name="psiginfo">Specifies information about the signature provider add-in.</param>
        /// <param name="XmlDsigStream">no description available</param>
        /// <returns>IPictureDisp</returns>
        [DispId(1610743808)]
        [MethodImpl(4096)]
        [return: ComAliasName("stdole.IPictureDisp")]
        [return: MarshalAs(28)]
        IPictureDisp GenerateSignatureLineImage([In] SignatureLineImage siglnimg, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo, [MarshalAs(25)] [In] object XmlDsigStream);

        /// <summary>
        /// Provides a signature provider add-in the opportunity to display the Signature Setup dialog box to the user. 
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.showsignaturesetup.aspx </remarks>
        /// <param name="ParentWindow">Contains the handle to the window containing the Signature Setup dialog box.</param>
        /// <param name="psigsetup">Specifies initial settings of the signature provider.</param>
        [DispId(1610743809)]
        [MethodImpl(4096)]
        void ShowSignatureSetup([MarshalAs(25)] [In] object ParentWindow, [MarshalAs(28)] [In] SignatureSetup psigsetup);

        /// <summary>
        /// Provides a signature provider add-in the opportunity to display the Signature dialog box to users, allowing them to specify their identity and then be authenticated.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.showsigningceremony.aspx </remarks>
        /// <param name="ParentWindow">Contains the handle to the window containing the Signature dialog box.</param>
        /// <param name="psigsetup">Specifies initial settings of the signature provider.</param>
        /// <param name="psiginfo">Specifies information about the signature provider.</param>
        [DispId(1610743810)]
        [MethodImpl(4096)]
        void ShowSigningCeremony([MarshalAs(25)] [In] object ParentWindow, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo);

        /// <summary>
        /// Used to sign the XMLDSIG template.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.signxmldsig.aspx </remarks>
        /// <param name="QueryContinue">Provides a way to query the host application for permission to continue the verification operation.</param>
        /// <param name="psigsetup">Specifies configuration information about a signature line.</param>
        /// <param name="psiginfo">Specifies information captured from the signing ceremony.</param>
        /// <param name="XmlDsigStream">Represents a steam of data containing XML, which represents an XMLDSIG object.</param>
        [DispId(1610743811)]
        [MethodImpl(4096)]
        void SignXmlDsig([MarshalAs(25)] [In] object QueryContinue, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo, [MarshalAs(25)] [In] object XmlDsigStream);

        /// <summary>
        /// Used to display a dialog box informing the user that the signing process has completed and providing additional functionality for the add-in.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.notifysignatureadded.aspx </remarks>
        /// <param name="ParentWindow">Allows the host application to obtain the handle to the window containing the displayed dialog box.</param>
        /// <param name="psigsetup">Contains initial settings of the signature provider.</param>
        /// <param name="psiginfo">Contains information about the signature provider add-in.</param>
        [DispId(1610743812)]
        [MethodImpl(4096)]
        void NotifySignatureAdded([MarshalAs(25)] [In] object ParentWindow, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo);

        /// <summary>
        /// Verifies a signature based on the signed state of the document and the legitimacy of the certificate used for signing.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.verifyxmldsig.aspx </remarks>
        /// <param name="QueryContinue">Provides a way to query the host application for permission to continue the verification operation.</param>
        /// <param name="psigsetup">Specifies configuration information about a signature line.</param>
        /// <param name="psiginfo">Specifies information captured from the signing ceremony.</param>
        /// <param name="XmlDsigStream">Represents a steam of data containing XML, which represents an XMLDSIG object</param>
        /// <param name="pcontverres">Specifies the status of the signature verification action.</param>
        /// <param name="pcertverres">Specifies the status of the signing certificate verification.</param>
        [DispId(1610743813)]
        [MethodImpl(4096)]
        void VerifyXmlDsig([MarshalAs(25)] [In] object QueryContinue, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo, [MarshalAs(25)] [In] object XmlDsigStream, [In] [Out] ref ContentVerificationResults pcontverres, [In] [Out] ref CertificateVerificationResults pcertverres);

        /// <summary>
        /// Provides a signature provider add-in the opportunity to display details about a signed signature line and display additional stored information such as a secure time-stamp.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.showsignaturedetails.aspx </remarks>
        /// <param name="ParentWindow">Contains the handle to the window containing the signature details.</param>
        /// <param name="psigsetup">Specifies initial settings of the signature provider.</param>
        /// <param name="psiginfo">Specifies information about the signed signature line.</param>
        /// <param name="XmlDsigStream">Represents a steam of data or binary large object of XML.</param>
        /// <param name="pcontverres">Contains a value representing the results of verificating the signature content.</param>
        /// <param name="pcertverres">Contains a value representing the results of verificating the signing certification.</param>
        [DispId(1610743814)]
        [MethodImpl(4096)]
        void ShowSignatureDetails([MarshalAs(25)] [In] object ParentWindow, [MarshalAs(28)] [In] SignatureSetup psigsetup, [MarshalAs(28)] [In] SignatureInfo psiginfo, [MarshalAs(25)] [In] object XmlDsigStream, [In] [Out] ref ContentVerificationResults pcontverres, [In] [Out] ref CertificateVerificationResults pcertverres);

        /// <summary>
        /// Queries the signature provider add-in for various details. 
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.getproviderdetail.aspx </remarks>
        /// <param name="sigprovdet">Contains an enumerated value representing the type of information to query the add-in for.</param>
        /// <returns>Object</returns>
        [DispId(1610743815)]
        [MethodImpl(4096)]
        [return: MarshalAs(27)]
        object GetProviderDetail([In] SignatureProviderDetail sigprovdet);

        /// <summary>
        /// Allows a signature provider add-in to create a hash value for the document that you can use to determine if the document contents were tampered with after digital signing.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/en-us/library/microsoft.office.core.signatureprovider.hashstream.aspx </remarks>
        /// <param name="QueryContinue">Provides a way to query the host application for permission to continue the hashing process.</param>
        /// <param name="Stream">Contains the data stream.</param>
        /// <returns>Array</returns>
        [DispId(1610743816)]
        [MethodImpl(4096)]
        [return: MarshalAs(29, SafeArraySubType = VarEnum.VT_UI1)]
        Array HashStream([MarshalAs(25)] [In] object QueryContinue, [MarshalAs(25)] [In] object Stream);
    }
}