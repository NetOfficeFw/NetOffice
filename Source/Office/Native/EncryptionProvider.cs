using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.Native
{
    /// <summary>
    /// Provides the methods for setting up permissions, applying the cryptography of the underlying encryption and decryption, and user authentication. 
    /// </summary>
    /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-object-office </remarks>
    [ComImport, Guid("000CD809-0000-0000-C000-000000000046"), TypeLibType(4160)]
    public interface EncryptionProvider
    {
        /// <summary>
        /// Displays information about the encryption of the current document. 
        /// </summary>
        /// <param name="encprovdet">Specifies the encryption information that you want.</param>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-getproviderdetail-method-office </remarks>
        /// <returns>Variant</returns>
        [DispId(1610743808)]
        [MethodImpl(4096)]
        [return: MarshalAs(27)]
        object GetProviderDetail([In] EncryptionProviderDetail encprovdet);

        /// <summary>
        /// Used by the EncryptionProvider object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-newsession-method-office </remarks>        
        /// <param name="ParentWindow">Specifies the window that is called to display the encryption settings</param>
        /// <returns>Long</returns>
        [DispId(1610743809)]
        [MethodImpl(4096)]
        int NewSession([MarshalAs(25)] [In] object ParentWindow);

        /// <summary>
        /// Used to determine whether the user has the proper permissions to open the encrypted document.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-authenticate-method-office </remarks>
        /// <param name="ParentWindow">Specifies the window that is called to display the encryption settings.</param>
        /// <param name="EncryptionData">Contains the encrypted data for the current document.</param>
        /// <param name="PermissionsMask">The user interface displayed by the encryption provider add-in.</param>
        /// <returns>Long</returns>
        [DispId(1610743810)]
        [MethodImpl(4096)]
        int Authenticate([MarshalAs(25)] [In] object ParentWindow, [MarshalAs(25)] [In] object EncryptionData, out uint PermissionsMask);

        /// <summary>
        /// Creates a second, working copy of the EncryptionProvider object's encryption session for a file that is about to be saved.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-clonesession-method-office </remarks>
        /// <param name="SessionHandle">The ID of the cloned session.</param>
        /// <returns>Long</returns>
        [DispId(1610743811)]
        [MethodImpl(4096)]
        int CloneSession([In] int SessionHandle);

        /// <summary>
        /// Ends the current encryption session.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-endsession-method-office </remarks>        
        /// <param name="SessionHandle">The ID of the current session.</param>
        [DispId(1610743812)]
        [MethodImpl(4096)]
        void EndSession([In] int SessionHandle);

        /// <summary>
        /// Saves an encrypted document.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-save-method-office </remarks>
        /// <param name="SessionHandle">The ID of the current session.</param>
        /// <param name="EncryptionData">Contains the encryption information.</param>
        /// <returns>Long</returns>
        [DispId(1610743813)]
        [MethodImpl(4096)]
        int Save([In] int SessionHandle, [MarshalAs(25)] [In] object EncryptionData);

        /// <summary>
        /// Encrypts and returns a stream of data for a document.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-encryptstream-method-office </remarks>
        /// <param name="SessionHandle">The ID of the current session.</param>
        /// <param name="StreamName">The name of the encrypted stream of document data.</param>
        /// <param name="UnencryptedStream">The data stream before encryption.</param>
        /// <param name="EncryptedStream">The data stream information after it has been encrypted.</param>
        [DispId(1610743814)]
        [MethodImpl(4096)]
        void EncryptStream([In] int SessionHandle, [MarshalAs(19)] [In] string StreamName, [MarshalAs(25)] [In] object UnencryptedStream, [MarshalAs(25)] [In] object EncryptedStream);

        /// <summary>
        /// Decrypts and returns a stream of encrypted data for a document.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-decryptstream-method-office </remarks>
        /// <param name="SessionHandle">The ID of the current session.</param>
        /// <param name="StreamName">The ID of the stream of data.</param>
        /// <param name="EncryptedStream">The encrypted data stream.</param>
        /// <param name="UnencryptedStream">The data stream before dencryption.</param>
        [DispId(1610743815)]
        [MethodImpl(4096)]
        void DecryptStream([In] int SessionHandle, [MarshalAs(19)] [In] string StreamName, [MarshalAs(25)] [In] object EncryptedStream, [MarshalAs(25)] [In] object UnencryptedStream);

        /// <summary>
        /// Used to display a dialog of the encryption settings for the current document.
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/encryptionprovider-showsettings-method-office </remarks>
        /// <param name="SessionHandle">The ID of the current session.</param>
        /// <param name="ParentWindow">Specifies the window that is called to display the encryption settings.</param>
        /// <param name="ReadOnly">Specifies whether you want the user to be able to change the encryption settings.</param>
        /// <param name="Remove">If True the encryption for a document will be removed during the next save operation.</param>
        [DispId(1610743816)]
        [MethodImpl(4096)]
        void ShowSettings([In] int SessionHandle, [MarshalAs(25)] [In] object ParentWindow, [In] bool ReadOnly, out bool Remove);
    }
}
